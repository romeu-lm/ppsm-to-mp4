import time
from pathlib import Path

import pythoncom
import win32com.client


# Office / PowerPoint constants
ppFixedFormatTypePDF = 2          # PpFixedFormatType
ppFixedFormatIntentPrint = 2      # PpFixedFormatIntent
ppSaveAsPDF = 32                  # PpSaveAsFileType

msoTrue = -1
msoFalse = 0

# MsoShapeType
msoMedia = 16
msoGroup = 6


def wait_for_file(out_file: Path, timeout_s: int = 120, min_bytes: int = 10_000) -> None:
    """Wait until file exists and has a reasonable size (PowerPoint sometimes returns before flush)."""
    t0 = time.time()
    while True:
        pythoncom.PumpWaitingMessages()

        if out_file.exists():
            try:
                if out_file.stat().st_size >= min_bytes:
                    return
            except OSError:
                pass

        if time.time() - t0 > timeout_s:
            raise TimeoutError(f"Timeout waiting for file to be created: {out_file}")

        time.sleep(0.5)


def _is_media_or_cameo_shape(sh) -> bool:
    """Best-effort detection: media shapes, or shapes exposing MediaFormat / Cameo."""
    try:
        if int(sh.Type) == msoMedia:
            return True
    except Exception:
        pass

    # Some shapes don't report Type reliably; MediaFormat throws if not media
    try:
        _ = sh.MediaFormat
        return True
    except Exception:
        pass

    # Newer PowerPoint: Cameo property may exist (throws if not)
    try:
        _ = sh.Cameo
        return True
    except Exception:
        pass

    return False


def _looks_like_webcam_overlay(sh, slide_w: float, slide_h: float) -> bool:
    """
    Heuristic: webcam overlay is usually a relatively small rectangle in bottom-right.
    Tune thresholds if needed.
    """
    try:
        left = float(sh.Left)
        top = float(sh.Top)
        w = float(sh.Width)
        h = float(sh.Height)
    except Exception:
        return False

    # bottom-right region
    in_corner = (left > 0.65 * slide_w) and (top > 0.65 * slide_h)

    # not huge
    smallish = (w < 0.50 * slide_w) and (h < 0.50 * slide_h)

    # also try name/alt-text hints (optional)
    name = ""
    alt = ""
    try:
        name = (sh.Name or "").lower()
    except Exception:
        pass
    try:
        alt = (sh.AlternativeText or "").lower()
    except Exception:
        pass

    hint = any(k in name for k in ("camera", "cameo", "webcam", "presenter")) or \
           any(k in alt for k in ("camera", "cameo", "webcam", "presenter"))

    # Require media/cameo OR explicit hint; don't delete random pictures.
    return in_corner and smallish and (_is_media_or_cameo_shape(sh) or hint)


def _delete_webcam_shapes_in_shapes(shapes, slide_w: float, slide_h: float) -> int:
    """Delete webcam-like shapes in a Shapes collection. Returns number deleted."""
    deleted = 0

    # IMPORTANT: iterate backwards when deleting
    for i in range(shapes.Count, 0, -1):
        sh = shapes.Item(i)
        try:
            # If it's a group, sometimes the webcam is the whole group; treat group as a unit
            is_group = False
            try:
                is_group = int(sh.Type) == msoGroup
            except Exception:
                pass

            if _looks_like_webcam_overlay(sh, slide_w, slide_h):
                sh.Delete()
                deleted += 1
                continue

            # If group but not detected as a whole, optionally scan inside
            if is_group:
                try:
                    gi = sh.GroupItems
                    for j in range(gi.Count, 0, -1):
                        inner = gi.Item(j)
                        if _looks_like_webcam_overlay(inner, slide_w, slide_h):
                            # delete the entire group (cleaner)
                            sh.Delete()
                            deleted += 1
                            break
                except Exception:
                    pass

        except Exception:
            # never fail export because one shape can't be inspected
            pass

    return deleted


def remove_webcam_overlay(pres) -> int:
    """
    Remove webcam/cameo overlay from:
      - all slides
      - slide masters and custom layouts (common place for Cameo)
    Returns total number of shapes deleted.
    """
    slide_w = float(pres.PageSetup.SlideWidth)
    slide_h = float(pres.PageSetup.SlideHeight)

    total_deleted = 0

    # Slides
    for s in range(1, pres.Slides.Count + 1):
        slide = pres.Slides.Item(s)
        total_deleted += _delete_webcam_shapes_in_shapes(slide.Shapes, slide_w, slide_h)

    # Masters / layouts (covers "on every slide" overlays)
    try:
        for d in range(1, pres.Designs.Count + 1):
            master = pres.Designs.Item(d).SlideMaster
            total_deleted += _delete_webcam_shapes_in_shapes(master.Shapes, slide_w, slide_h)

            try:
                for l in range(1, master.CustomLayouts.Count + 1):
                    layout = master.CustomLayouts.Item(l)
                    total_deleted += _delete_webcam_shapes_in_shapes(layout.Shapes, slide_w, slide_h)
            except Exception:
                pass
    except Exception:
        pass

    return total_deleted


def export_folder_ppsm_to_pdf_no_webcam(
    input_dir: str,
    output_dir: str,
    print_hidden_slides: bool = False,
    print_quality: bool = True,
):
    in_path = Path(input_dir)
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    ppt = None

    try:
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        ppt.Visible = True
        ppt.DisplayAlerts = 0

        # Disable macros on open (avoids prompts for PPSM)
        # 3 = msoAutomationSecurityForceDisable
        try:
            ppt.AutomationSecurity = 3
        except Exception:
            pass

        files = sorted(in_path.glob("*.ppsm"))
        if not files:
            print("No .ppsm files found.")
            return

        for ppsm in files:
            out_pdf = out_path / (ppsm.stem + ".pdf")
            print(f"Exporting: {ppsm.name} -> {out_pdf.name}")

            # Open(FileName, ReadOnly, Untitled, WithWindow)
            pres = ppt.Presentations.Open(str(ppsm), True, False, False)

            try:
                deleted = remove_webcam_overlay(pres)
                if deleted:
                    print(f"  removed {deleted} webcam/cameo shape(s)")

                intent = ppFixedFormatIntentPrint if print_quality else 1  # 1=screen
                pres.ExportAsFixedFormat(
                    str(out_pdf),
                    ppFixedFormatTypePDF,
                    intent,
                    msoFalse,   # FrameSlides
                    PrintHiddenSlides=msoTrue if print_hidden_slides else msoFalse,
                )
            except Exception:
                # Fallback
                pres.SaveAs(str(out_pdf), ppSaveAsPDF)

            wait_for_file(out_pdf)

            # prevent "save changes?"" prompts
            try:
                pres.Saved = True
            except Exception:
                pass

            pres.Close()

        print("Done.")

    finally:
        if ppt is not None:
            try:
                ppt.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    base_dir = Path(__file__).resolve().parent

    export_folder_ppsm_to_pdf_no_webcam(
        input_dir=base_dir,
        output_dir=base_dir / "Slides",
        print_hidden_slides=False,
        print_quality=True,
    )
