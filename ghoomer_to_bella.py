#! python3
"""
oom_anim_bake — Grasshopper frame bake → .bsz per frame + optional render.cmd / render.sh

Bakes each frame of a defined GH definition (Slider + GEO_OUT + LAYER_OUT) into
Rhino, exports Bella .bsz, then removes baked geometry. The dialog loads the
slider’s min/max from the .gh and lets you pick a start/end frame (subset).
Optional: write a small render loop script for bella_cli.

Eto dialog (default): set paths and Grasshopper nicknames without
editing this file. Every TableRow cell is wrapped in TableCell(…)
so the Eto → .NET bridge does not mangle layout.

Run with no arguments for the dialog. For scripts/automation, use --no-gui and
rely on defaults (same as the old script) or set env / extend argparse later.
"""

from __future__ import annotations

import os
import sys
import uuid
import System
import System.Reflection
import argparse

# --- Bella CLI (from Bella plugin; Rhino only) ------------------------------------
import clr  # type: ignore
import bella  # type: ignore
import Rhino  # type: ignore

_dotnet_dll = System.Reflection.Assembly.GetAssembly(clr.GetClrType(bella.bella)).Location
_plugin_dir = os.path.dirname(_dotnet_dll)

if Rhino.Runtime.HostUtils.RunningOnWindows:
    _BELLA_CLI = os.path.join(_plugin_dir, "bella", "bella_cli.exe")
else:
    _BELLA_CLI = os.path.join(
        _plugin_dir, "bella", "bella_cli.app", "Contents", "MacOS", "bella_cli"
    )

# --- Defaults (no-dialog / --no-gui) match the original oom_anim_bake script -------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_DEFAULT_GH = os.path.join(_SCRIPT_DIR, "oom_anim.gh")
_DEFAULT_OUT_DIR = os.path.join(_SCRIPT_DIR, "bsz")
_DEFAULT_OUT_NAME = "oom_bake"
_DEFAULT_SLIDER = "FRAME"
_DEFAULT_GEO = "GEO_OUT"
_DEFAULT_LAYER = "LAYER_OUT"


def _eto_available() -> bool:
    try:
        import Rhino  # noqa: F401
        import Eto.Forms  # noqa: F401
        import Eto.Drawing  # noqa: F401
    except Exception:
        return False
    return True


def _get_slider_frame_range(gh_path: str, slider_nick: str) -> tuple[int, int] | None:
    """
    Return (min, max) integer frame indices from a Grasshopper number slider
    in the given .gh, or None if the file or nickname is invalid.
    """
    if not gh_path or not os.path.isfile(gh_path):
        return None
    nick = (slider_nick or "").strip() or _DEFAULT_SLIDER
    import Grasshopper as gh  # type: ignore

    gh_doc_io = gh.Kernel.GH_DocumentIO()
    if not gh_doc_io.Open(gh_path):
        return None
    gh_doc = gh_doc_io.Document
    try:
        slider = next(
            (obj for obj in gh_doc.Objects if obj.NickName == nick),
            None,
        )
        if not slider:
            return None
        try:
            inner = slider.Slider
        except Exception:
            return None
        if inner is None:
            return None
        lo = int(System.Convert.ToInt32(inner.Minimum))
        hi = int(System.Convert.ToInt32(inner.Maximum))
        if hi < lo:
            lo, hi = hi, lo
        return (lo, hi)
    finally:
        try:
            gh_doc.Dispose()
        except Exception:
            pass


def _activate_first_perspective_view(doc) -> bool:
    """Set `doc.Views.ActiveView` to the first non-parallel (perspective) viewport, if any."""
    try:
        from Rhino.Display import RhinoPageView  # type: ignore
    except Exception:
        RhinoPageView = None  # type: ignore
    try:
        for view in doc.Views:
            try:
                if RhinoPageView is not None and isinstance(view, RhinoPageView):
                    continue
                vp = view.ActiveViewport
                if vp is None:
                    continue
                if not vp.IsParallelProjection:
                    doc.Views.ActiveView = view
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _render_source_uses_active_viewport_only(doc) -> bool:
    """
    True when **Document → Render** is set to use the *current* / active viewport for rendering,
    as opposed to a specific view or named snapshot. In that case we prefer the first
    perspective view for Bella, not whatever panel was last clicked.
    """
    import Rhino  # type: ignore

    try:
        src = doc.RenderSettings.RenderSource
    except Exception:
        return True
    try:
        R = Rhino.Render.RenderSettingsRenderingSources
        for name in (
            "ActiveView",
            "ActiveViewport",
            "DefaultView",
        ):
            if hasattr(R, name) and src == getattr(R, name):
                return True
    except Exception:
        pass
    return False


def _run_bella_export_with_render_view(doc, cmd: str) -> None:
    """
    BellaExport follows the active viewport. When **Render view source** is a specific
    view, `RenderSourceView` + `GetViewInfo()` applies it; if render uses “active viewport”
    or the API is missing, the first *perspective* view is activated so a stray Top/Front
    selection does not drive the export.
    """
    import Rhino  # type: ignore

    rsv = None
    try:
        rsv = Rhino.Render.RenderSourceView(doc)
    except Exception:
        rsv = None

    if rsv is not None:
        try:
            try:
                rsv.GetViewInfo()
            except Exception:
                _activate_first_perspective_view(doc)
            else:
                if _render_source_uses_active_viewport_only(doc):
                    _activate_first_perspective_view(doc)
            Rhino.RhinoApp.RunScript(cmd, False)
        finally:
            try:
                rsv.Dispose()
            except Exception:
                pass
        return

    _activate_first_perspective_view(doc)
    Rhino.RhinoApp.RunScript(cmd, False)


def _show_bake_dialog() -> dict | None:
    """Eto form; returns a settings dict, or None if cancelled."""
    import Rhino  # noqa: F401
    import Eto.Forms as _ef
    from Eto.Drawing import Padding, Size
    from Eto.Forms import (
        Button,
        CheckBox,
        Dialog,
        FileFilter,
        FilePicker,
        HorizontalAlignment,
        Label,
        LinkButton,
        MessageBox,
        Orientation,
        StackLayout,
        StackLayoutItem,
        TableCell,
        TableLayout,
        TableRow,
        TextAlignment,
        TextBox,
        NumericUpDown,
    )

    FORM_LABEL_WIDTH = 130
    DLG_DEFAULT = Size(820, 560)
    DLG_MIN = Size(640, 480)
    CONTROL_MIN_WIDTH = 400
    # Before .gh is read, NumericUpDown must not use Min=Max=0 (typing is locked).
    _FRAME_NU_MAX = 9_999_999.0

    def _mk_label(text, fixed_width=False):
        lbl = Label()
        lbl.Text = text
        if fixed_width:
            lbl.Width = FORM_LABEL_WIDTH
        return lbl

    def _mk_button(text, handler):
        btn = Button()
        btn.Text = text
        btn.Click += handler
        return btn

    def _min_w(ctrl, w):
        try:
            ctrl.MinimumSize = Size(w, 0)
        except Exception:
            pass

    def _stretched_stack_item(layout_control, expand_along_stack: bool) -> StackLayoutItem:
        """Vertical StackLayout: the bool is vertical expansion, not width. For width,
        the stack and each item use HorizontalContentAlignment=Stretch and per-item
        HorizontalAlignment=Stretch; otherwise children stay at preferred size and look
        clipped on resize."""
        item = StackLayoutItem(layout_control, expand_along_stack)
        try:
            item.HorizontalAlignment = HorizontalAlignment.Stretch
        except Exception:
            pass
        return item

    class BakeDialog(Dialog):
        def __init__(self):
            super().__init__()
            self.Title = "oom anim bake — Grasshopper → .bsz"
            self.ClientSize = DLG_DEFAULT
            try:
                self.MinimumSize = DLG_MIN
            except Exception:
                pass
            self.Padding = Padding(10)
            self.Resizable = True
            self._result: dict | None = None
            self._last_gh_nick: tuple[str, str] | None = None

            intro = _mk_label(
                "Bakes one .bsz per Slider (default nickname FRAME) from a Grasshopper file.\n"
                "Requires two data outputs (geometry + layer index); nicknames GEO_OUT and LAYER_OUT."
            )
            try:
                intro.Wrap = _ef.WrapMode.Word
            except Exception:
                pass

            heading = _mk_label(
                "Take a Grasshopper file driven by a single  animation slider and bake to .bsz files"
            )
            try:
                heading.Wrap = _ef.WrapMode.Word
            except Exception:
                pass
            try:
                intro.TextAlignment = TextAlignment.Left
                heading.TextAlignment = TextAlignment.Left
            except Exception:
                pass

            self.gh_picker = FilePicker()
            f = FileFilter()
            f.Name = "Grasshopper (*.gh)"
            f.Extensions = [".gh"]
            self.gh_picker.Filters.Add(f)
            if os.path.isfile(_DEFAULT_GH):
                self.gh_picker.FilePath = _DEFAULT_GH
            _min_w(self.gh_picker, CONTROL_MIN_WIDTH)

            self.out_dir = TextBox()
            self.out_dir.Text = _DEFAULT_OUT_DIR
            _min_w(self.out_dir, CONTROL_MIN_WIDTH)
            self.out_dir.ToolTip = "Folder for BellaExport .bsz output (created if needed)."

            self.out_name = TextBox()
            self.out_name.Text = _DEFAULT_OUT_NAME
            _min_w(self.out_name, CONTROL_MIN_WIDTH)

            self.nick_slider = TextBox()
            self.nick_slider.Text = _DEFAULT_SLIDER
            _min_w(self.nick_slider, CONTROL_MIN_WIDTH)

            self.nick_geo = TextBox()
            self.nick_geo.Text = _DEFAULT_GEO
            _min_w(self.nick_geo, CONTROL_MIN_WIDTH)

            self.nick_layer = TextBox()
            self.nick_layer.Text = _DEFAULT_LAYER
            _min_w(self.nick_layer, CONTROL_MIN_WIDTH)

            self.frame_start = NumericUpDown()
            self.frame_end = NumericUpDown()
            for nu in (self.frame_start, self.frame_end):
                try:
                    nu.DecimalPlaces = 0
                except Exception:
                    try:
                        nu.MaxDecimalPlaces = 0
                    except Exception:
                        pass
                nu.MinValue = 0.0
                nu.MaxValue = _FRAME_NU_MAX
                nu.Value = 0.0
                nu.Increment = 1.0
                _min_w(nu, 180)

            self.cb_frame_tie_gh = CheckBox()
            self.cb_frame_tie_gh.Text = "Clamp to .gh slider min/max (auto on path change)"
            self.cb_frame_tie_gh.ToolTip = (
                "On: the .gh is read; spinners match the slider’s min/max and update when the "
                "path or nickname changes. Off (expert): type any values without querying first; "
                "bake still clamps to the real slider. Use “Load range from .gh” to copy suggested start/end."
            )
            self.cb_frame_tie_gh.Checked = True

            self.cb_zero_is_full = CheckBox()
            self.cb_zero_is_full.Text = "0+0 = full .gh range (start & end both zero)"
            self.cb_zero_is_full.ToolTip = (
                "When both numbers are 0, bake the whole animation range from the .gh. "
                "Turn off to bake only frame 0 (start=0, end=0)."
            )
            self.cb_zero_is_full.Checked = True

            self.frame_start.ToolTip = (
                "First frame (inclusive). With “Clamp” off, you can type before opening the .gh. "
                "With “0+0 = full” on, 0+0 bakes the full slider range."
            )
            self.frame_end.ToolTip = (
                "Last frame (inclusive). With “Clamp” off, you can type before opening the .gh. "
                "With “0+0 = full” on, 0+0 bakes the full slider range."
            )

            self.btn_frame_sync = _mk_button("Load range from .gh", self._on_frame_sync)

            self.write_render = CheckBox()
            # CheckBox does not word-wrap on all platforms; keep the caption to one line.
            self.write_render.Text = "Write bella render script next to this .py (render.cmd / .sh)"
            self.write_render.ToolTip = (
                "Optional local loop: bella_cli -i bsz, -o png only. "
                "Set resolution, parseFragment, etc. on the farm or edit the file."
            )
            self.write_render.Checked = True

            def _cell(control, scale=False):
                return TableCell(control, scale)

            form = TableLayout()
            form.Padding = Padding(10)
            form.Spacing = Size(6, 6)

            def _row(lbl, ctrl):
                form.Rows.Add(
                    TableRow(
                        _cell(_mk_label(lbl, True)),
                        _cell(ctrl, True),
                    )
                )

            # One-column table so intro lines use the full content width, not a single
            # column of the two-column form below.
            info = TableLayout()
            info.Spacing = Size(6, 6)
            # Match horizontal inset of `form` (dialog padding + form padding).
            info.Padding = Padding(10, 0, 10, 4)
            info.Rows.Add(TableRow(_cell(heading, True)))
            info.Rows.Add(TableRow(_cell(intro, True)))

            _row("Grasshopper file", self.gh_picker)
            _row("Output folder", self.out_dir)
            _row("Base name (.bsz)", self.out_name)
            _row("Slider nickname", self.nick_slider)
            _row("Start frame (incl.)", self.frame_start)
            _row("End frame (incl.)", self.frame_end)
            form.Rows.Add(TableRow(_cell(None, True), _cell(self.cb_frame_tie_gh, True)))
            form.Rows.Add(TableRow(_cell(None, True), _cell(self.cb_zero_is_full, True)))
            form.Rows.Add(
                TableRow(
                    _cell(_mk_label(" ", True)),
                    _cell(self.btn_frame_sync, True),
                )
            )
            _row("GEO output nick", self.nick_geo)
            _row("LAYER output nick", self.nick_layer)
            form.Rows.Add(TableRow(_cell(None, True), _cell(self.write_render, True)))

            spacer = TableRow()
            spacer.ScaleHeight = True
            form.Rows.Add(spacer)

            self.run_btn = _mk_button("Bake", self._on_run)
            self.DefaultButton = self.run_btn
            self.cancel_btn = _mk_button("Cancel", self._on_cancel)
            self.AbortButton = self.cancel_btn

            self.test_bake_link = LinkButton()
            self.test_bake_link.Text = "dev · test bake"
            self.test_bake_link.ToolTip = (
                "Dev: bake only the first frame of your range, one .bsz, "
                "leave geometry in the document (you delete it)."
            )
            self.test_bake_link.Click += self._on_test_bake

            btn_bar = TableLayout()
            btn_bar.Spacing = Size(6, 0)
            btn_bar.Rows.Add(
                TableRow(
                    _cell(self.test_bake_link, False),
                    _cell(None, True),
                    _cell(self.cancel_btn, False),
                    _cell(self.run_btn, False),
                )
            )
            form.Rows.Add(TableRow(_cell(None), _cell(btn_bar, True)))

            root = StackLayout()
            root.Orientation = Orientation.Vertical
            # StackLayout.Spacing is an int (gap between children), not Eto.Drawing.Size.
            root.Spacing = 0
            # Children must stretch to full content width, or they keep preferred
            # width and clip / show dead margin when the user resizes the dialog.
            try:
                root.HorizontalContentAlignment = HorizontalAlignment.Stretch
            except Exception:
                pass
            # In a vertical stack, the StackLayoutItem bool is vertical "expand" only:
            # keep the intro to natural height; let the two-column form take the rest
            # (incl. flexible spacer) so the button bar stays usable.
            root.Items.Add(_stretched_stack_item(info, False))
            root.Items.Add(_stretched_stack_item(form, True))
            self.Content = root

            def _on_nick_or_gh(_s, _e):
                self._sync_frame_range_from_gh(from_button=False)

            try:
                self.nick_slider.TextChanged += _on_nick_or_gh
            except Exception:
                pass
            try:
                self.gh_picker.ValueChanged += _on_nick_or_gh
            except Exception:
                pass
            try:
                self.cb_frame_tie_gh.CheckedChanged += self._on_frame_tie_changed
            except Exception:
                pass
            self._sync_frame_range_from_gh(from_button=False)

        def _set_wide_frame_bounds(self) -> None:
            for nu in (self.frame_start, self.frame_end):
                try:
                    nu.MinValue = 0.0
                    nu.MaxValue = _FRAME_NU_MAX
                except Exception:
                    pass

        def _on_frame_tie_changed(self, sender, e):
            if self.cb_frame_tie_gh.Checked:
                self._last_gh_nick = None
                self._sync_frame_range_from_gh(from_button=False)
            else:
                self._set_wide_frame_bounds()

        def _on_frame_sync(self, sender, e):
            self._sync_frame_range_from_gh(from_button=True)

        def _sync_frame_range_from_gh(self, from_button: bool) -> None:
            """Set start/end from the FRAME slider in the chosen .gh.
            * Auto (from_button False): only when “Clamp to .gh” is on.
            * Load button (from_button True): always reads .gh; copies suggested start/end;
              narrows min/max only when “Clamp to .gh” is on.
            """
            if not from_button and not self.cb_frame_tie_gh.Checked:
                return
            gh = (self.gh_picker.FilePath or "").strip()
            nick = (self.nick_slider.Text or "").strip() or _DEFAULT_SLIDER
            cur_key = (gh, nick)
            rng = _get_slider_frame_range(gh, nick)
            if not rng:
                return
            lo, hi = rng
            try:
                lo_d = float(lo)
                hi_d = float(hi)
            except Exception:
                return
            tie = self.cb_frame_tie_gh.Checked
            if tie:
                for nu in (self.frame_start, self.frame_end):
                    try:
                        nu.MinValue = lo_d
                        nu.MaxValue = hi_d
                    except Exception:
                        pass
                if self._last_gh_nick != cur_key:
                    self._last_gh_nick = cur_key
                    self.frame_start.Value = lo_d
                    self.frame_end.Value = hi_d
                else:
                    try:
                        if self.frame_start.Value < lo_d or self.frame_start.Value > hi_d:
                            self.frame_start.Value = lo_d
                        if self.frame_end.Value < lo_d or self.frame_end.Value > hi_d:
                            self.frame_end.Value = hi_d
                    except Exception:
                        self.frame_start.Value = lo_d
                        self.frame_end.Value = hi_d
            else:
                if self._last_gh_nick != cur_key:
                    self._last_gh_nick = cur_key
                self.frame_start.Value = lo_d
                self.frame_end.Value = hi_d

        def _bad(self, msg: str, title: str) -> None:
            try:
                MessageBox.Show(self, msg, title, _ef.MessageBoxType.Warning)
            except Exception:
                try:
                    MessageBox.Show(msg, title)
                except Exception:
                    print(msg, file=sys.stderr)

        def _collect_settings(self, test_bake: bool) -> bool:
            gh = (self.gh_picker.FilePath or "").strip()
            if not gh or not os.path.isfile(gh):
                self._bad("Choose a valid .gh file.", "Grasshopper")
                return False
            odir = (self.out_dir.Text or "").strip()
            if not odir:
                self._bad("Set an output folder for .bsz files.", "Output")
                return False
            oname = (self.out_name.Text or "").strip() or "bake"
            lo_hi = _get_slider_frame_range(gh, (self.nick_slider.Text or "").strip() or _DEFAULT_SLIDER)
            if not lo_hi:
                self._bad(
                    "Could not read the animation slider from the .gh. "
                    "Check the file and the Slider nickname, then use “Load range from .gh” "
                    "or try again.",
                    "Frame range",
                )
                return False
            min_v, max_v = lo_hi
            try:
                fs = int(self.frame_start.Value)
                fe = int(self.frame_end.Value)
            except Exception:
                fs, fe = 0, 0

            if self.cb_zero_is_full.Checked and fs == 0 and fe == 0:
                fsa, fea = None, None
            elif self.cb_frame_tie_gh.Checked:
                fsa = max(min_v, min(fs, max_v))
                fea = max(min_v, min(fe, max_v))
                if fsa > fea:
                    fsa, fea = fea, fsa
            else:
                fsa, fea = fs, fe

            self._result = {
                "gh_path": os.path.normpath(gh),
                "out_dir": os.path.normpath(odir),
                "out_name": oname,
                "slider_nick": (self.nick_slider.Text or "").strip() or _DEFAULT_SLIDER,
                "geo_nick": (self.nick_geo.Text or "").strip() or _DEFAULT_GEO,
                "layer_nick": (self.nick_layer.Text or "").strip() or _DEFAULT_LAYER,
                "frame_start": fsa,
                "frame_end": fea,
                "write_render": bool(self.write_render.Checked),
                "test_bake": test_bake,
            }
            return True

        def _on_run(self, sender, e):
            if not self._collect_settings(test_bake=False):
                return
            self.Close()

        def _on_test_bake(self, sender, e):
            if not self._collect_settings(test_bake=True):
                return
            self.Close()

        def _on_cancel(self, sender, e):
            self._result = None
            self.Close()

    import Rhino.UI  # type: ignore  # RhinoEtoApp for modal parent

    dlg = BakeDialog()
    try:
        owner = Rhino.UI.RhinoEtoApp.MainWindow
    except Exception:
        owner = None
    if owner is not None:
        dlg.ShowModal(owner)
    else:
        dlg.ShowModal()
    return dlg._result


def _default_settings() -> dict:
    return {
        "gh_path": _DEFAULT_GH,
        "out_dir": _DEFAULT_OUT_DIR,
        "out_name": _DEFAULT_OUT_NAME,
        "slider_nick": _DEFAULT_SLIDER,
        "geo_nick": _DEFAULT_GEO,
        "layer_nick": _DEFAULT_LAYER,
        "frame_start": None,
        "frame_end": None,
        "write_render": True,
        "test_bake": False,
    }


def _relpath_posix(a: str, b: str) -> str:
    """Path from b to a, with forward slashes (for -i: / -o: in .cmd and .sh)."""
    r = os.path.relpath(a, b)
    if r == ".":
        return "."
    return r.replace(os.sep, "/")


def _write_render_scripts(
    *,
    bella_cli: str,
    script_dir: str,
    out_dir: str,
    out_name: str,
    start_frame: int,
    end_frame: int,
) -> str:
    """
    Windows: write ``render.cmd`` (``cmd.exe /c``) with ``.\\bella\\bella_cli.exe`` and
    ``%%i`` in the for-loop, ``-i``/``-o`` only (no -res or -pf; use the farm or edit).

    macOS / Linux: write ``render.sh`` (bash) with an absolute path to bella_cli.
    Returns the path to the file written.
    """
    script_dir = os.path.abspath(script_dir)
    out_dir = os.path.abspath(out_dir)
    png_dir = os.path.join(script_dir, "png")
    if not os.path.isdir(png_dir):
        os.makedirs(png_dir, exist_ok=True)
    try:
        out_rel = _relpath_posix(out_dir, script_dir)
        png_rel = _relpath_posix(png_dir, script_dir)
    except ValueError:
        out_rel = out_dir.replace(os.sep, "/")
        png_rel = "png"

    if Rhino.Runtime.HostUtils.RunningOnWindows:
        path = os.path.join(script_dir, "render.cmd")
        in_arg = f"{out_rel}/{out_name}_%%i.bsz" if out_rel != "." else f"{out_name}_%%i.bsz"
        out_arg = f"{png_rel}/%%i.png" if png_rel != "." else "%%i.png"
        # Same shape as: .\bella\bella_cli.exe -i:bsx/goomer_%%i.bsx -o:png/%%i.png …
        with open(path, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(
                "@REM Run from a folder that contains bella\\bella_cli.exe (Bella for Rhino plugin layout), "
                f"or edit the path below. Frames {start_frame}..{end_frame}. -res, -pf, etc. on the farm or here.\n"
            )
            f.write(f"for /l %%i in ({start_frame},1,{end_frame}) do (\n")
            f.write(f"  .\\bella\\bella_cli.exe -i:{in_arg} -o:{out_arg}\n")
            f.write(")\n")
    else:
        path = os.path.join(script_dir, "render.sh")
        in_arg = f"{out_rel}/{out_name}_$i.bsz" if out_rel != "." else f"{out_name}_$i.bsz"
        out_arg = f"{png_rel}/$i.png" if png_rel != "." else "$i.png"
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write("#!/usr/bin/env bash\n")
            f.write(
                f"# Frames {start_frame}..{end_frame} — `chmod +x render.sh` then run from this directory.\n"
            )
            f.write(
                f"# -res, parseFragment (-pf), etc. are for your render farm or add below.\n"
            )
            f.write("set -e\n")
            f.write(f"for i in {{{start_frame}..{end_frame}}}; do\n")
            f.write(f'  "{bella_cli}" -i:"{in_arg}" -o:"{out_arg}"\n')
            f.write("done\n")
    return path


def run_oom_bake(s: dict) -> tuple:
    """
    Run the grasshopper → bake → BellaExport → delete loop (or test bake: one frame, keep geometry).
    Returns (ok: bool, message: str) for the UI or batch mode.
    """
    import Grasshopper as gh  # type: ignore

    gh_file_path = s["gh_path"]
    output_folder = s["out_dir"]
    outname = s["out_name"]
    slider_name = s["slider_nick"]
    geo_node_name = s["geo_nick"]
    layer_node_name = s["layer_nick"]
    test_bake = bool(s.get("test_bake", False))
    write_render = bool(s.get("write_render", True)) and not test_bake

    if not os.path.isfile(gh_file_path):
        return (False, f"Grasshopper file not found: {gh_file_path!r}")

    if not os.path.isdir(output_folder):
        try:
            os.makedirs(output_folder, exist_ok=True)
        except OSError as e:
            return (False, f"Cannot create output folder: {e}")

    active_doc = Rhino.RhinoDoc.ActiveDoc
    if not active_doc:
        return (False, "No active Rhino document.")

    gh_doc_io = gh.Kernel.GH_DocumentIO()
    if not gh_doc_io.Open(gh_file_path):
        return (False, f"Failed to open Grasshopper: {gh_file_path!r}")

    gh_doc = gh_doc_io.Document
    try:
        gh_doc.Enabled = True

        slider = next(
            (obj for obj in gh_doc.Objects if obj.NickName == slider_name), None
        )
        geo_node = next(
            (obj for obj in gh_doc.Objects if obj.NickName == geo_node_name), None
        )
        layer_node = next(
            (obj for obj in gh_doc.Objects if obj.NickName == layer_node_name), None
        )

        if not (slider and geo_node and layer_node):
            return (
                False,
                f"Need GH nicknames: Slider={slider_name!r}, geo={geo_node_name!r}, "
                f"layer={layer_node_name!r} — at least one node missing.",
            )

        print(
            f"oom_anim_bake: found nodes {slider_name!r}, {geo_node_name!r}, {layer_node_name!r}."
        )
        inner_slider = slider.Slider
        lo = int(System.Convert.ToInt32(inner_slider.Minimum))
        hi = int(System.Convert.ToInt32(inner_slider.Maximum))
        if hi < lo:
            lo, hi = hi, lo

        fs = s.get("frame_start")
        fe = s.get("frame_end")
        if fs is None:
            fs = lo
        if fe is None:
            fe = hi
        start_frame = int(fs)
        end_frame = int(fe)
        start_frame = max(lo, min(start_frame, hi))
        end_frame = max(lo, min(end_frame, hi))
        if start_frame > end_frame:
            start_frame, end_frame = end_frame, start_frame
        if test_bake:
            end_frame = start_frame
        n_frames = end_frame - start_frame + 1
        out_dir_norm = output_folder

        for i in range(start_frame, end_frame + 1):
            slider.SetSliderValue(System.Decimal(i))
            slider.ExpireSolution(False)
            gh_doc.NewSolution(True)

            baked_ids = []
            item_counter = 0

            geo_tree = geo_node.VolatileData
            layer_tree = layer_node.VolatileData

            for branch_idx in range(geo_tree.PathCount):
                path = geo_tree.get_Path(branch_idx)
                branch = geo_tree.get_Branch(path)

                target_layer_index = 0
                layer_branch = layer_tree.get_Branch(path)
                if layer_branch and layer_branch.Count > 0:
                    try:
                        target_layer_index = int(layer_branch[0].ScriptVariable())
                    except Exception:
                        print(
                            f"Warning: could not parse layer index for path {path}"
                        )

                for item in branch:
                    if item is None:
                        continue
                    raw_geom = item.ScriptVariable()
                    attr = Rhino.DocObjects.ObjectAttributes()
                    attr.LayerIndex = target_layer_index
                    attr.MaterialSource = (
                        Rhino.DocObjects.ObjectMaterialSource.MaterialFromLayer
                    )

                    if raw_geom:
                        consistent_str = f"per_mesh__{item_counter}"
                        new_guid_str = str(
                            uuid.uuid5(uuid.NAMESPACE_DNS, consistent_str)
                        )
                        attr.ObjectId = System.Guid(new_guid_str)
                        if i == start_frame:
                            _ = f"id_xform_{new_guid_str.replace('-', '_')}"

                        obj_id = active_doc.Objects.Add(raw_geom, attr)
                        if obj_id != System.Guid.Empty:
                            baked_ids.append(obj_id)
                        item_counter += 1

            if baked_ids:
                cmd = (
                    f"-_BellaExport _dir={out_dir_norm} _name={outname}_{i} "
                    f"_ext=.bsz _startGui=no _Enter"
                )
                _run_bella_export_with_render_view(active_doc, cmd)

                if not test_bake:
                    for b_id in baked_ids:
                        active_doc.Objects.Delete(b_id, True)

        if test_bake:
            msg_tail = (
                f"Test bake: exported one .bsz (frame {start_frame}) to {out_dir_norm!r}. "
                f"Baked geometry was left in the Rhino document."
            )
        else:
            msg_tail = f"Exported {n_frames} frame(s) (frames {start_frame}..{end_frame}) to {out_dir_norm!r}."
        if write_render and os.path.isfile(_BELLA_CLI):
            rpath = _write_render_scripts(
                bella_cli=_BELLA_CLI,
                script_dir=_SCRIPT_DIR,
                out_dir=out_dir_norm,
                out_name=outname,
                start_frame=start_frame,
                end_frame=end_frame,
            )
            print(f"oom_anim_bake: wrote {rpath!r}")
            msg_tail += f" Render script: {rpath!r}."
        elif write_render and not os.path.isfile(_BELLA_CLI):
            print(
                f"oom_anim_bake: bella_cli not at {_BELLA_CLI!r} — skip render script.",
                file=sys.stderr,
            )
        return (True, f"Done. {msg_tail}")
    finally:
        try:
            gh_doc.Dispose()
        except Exception:
            pass


def main() -> int:
    p = argparse.ArgumentParser(
        description="Bake GH animation frames to .bsz (Eto dialog by default).",
    )
    p.add_argument(
        "--no-gui",
        action="store_true",
        help="Run immediately with default paths and nicknames (no dialog).",
    )
    args, _unk = p.parse_known_args()
    s = _default_settings() if args.no_gui else None
    if s is None:
        if not _eto_available():
            print(
                "Eto not available; using defaults. For the dialog, run in Rhino 8+ "
                "with Eto, or use --no-gui.",
                file=sys.stderr,
            )
            s = _default_settings()
        else:
            s = _show_bake_dialog()
            if s is None:
                print("Cancelled.")
                # Return 0 so __main__ does not raise SystemExit: Rhino’s ScriptEditor
                # shows SystemExit(1) / tracebacks like a real failure, even for Cancel.
                return 0
    try:
        ok, msg = run_oom_bake(s)
    except Exception as e:  # noqa: BLE001
        ok, msg = False, f"{type(e).__name__}: {e}"
    if ok:
        print(msg)
        if _eto_available() and s is not None and not args.no_gui:
            try:
                import Eto.Forms as _ef2
                from Eto.Forms import MessageBox

                try:
                    MessageBox.Show(
                        msg,
                        "oom anim bake",
                        _ef2.MessageBoxType.Information,
                    )
                except Exception:
                    MessageBox.Show(msg, "oom anim bake")
            except Exception:
                pass
        return 0
    print(msg, file=sys.stderr)
    if _eto_available() and s is not None and not args.no_gui:
        try:
            import Eto.Forms as _ef2
            from Eto.Forms import MessageBox

            try:
                MessageBox.Show(
                    msg,
                    "oom anim bake — error",
                    _ef2.MessageBoxType.Error,
                )
            except Exception:
                MessageBox.Show(msg, "oom anim bake")
        except Exception:
            pass
    return 1


if __name__ == "__main__":
    # Don’t raise SystemExit(0): Rhino’s ScriptEditor often highlights it like an error.
    _rc = int(main() or 0)
    if _rc != 0:
        raise SystemExit(_rc)