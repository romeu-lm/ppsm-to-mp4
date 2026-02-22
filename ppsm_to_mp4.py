import time
from pathlib import Path

import pythoncom
import win32com.client


# PpMediaTaskStatus (PowerPoint)
PP_NONE = 0
PP_IN_PROGRESS = 1
PP_QUEUED = 2
PP_DONE = 3
PP_FAILED = 4


def wait_for_video(pres, out_mp4: Path, timeout_s: int = 60 * 60) -> None:
    """
    Wait until PowerPoint finishes CreateVideo.
    Also wait until the MP4 file exists and is non-trivial in size.
    """
    t0 = time.time()
    last_status = None

    while True:
        pythoncom.PumpWaitingMessages()

        status = int(pres.CreateVideoStatus)

        if status != last_status:
            print(f"  status = {status}")
            last_status = status

        if status == PP_FAILED:
            raise RuntimeError(f"PowerPoint export FAILED for: {out_mp4.name}")

        if status == PP_DONE:
            break

        if time.time() - t0 > timeout_s:
            raise TimeoutError(f"Timeout waiting for export: {out_mp4.name}")

        time.sleep(1)

    t1 = time.time()
    while True:
        if out_mp4.exists():
            try:
                if out_mp4.stat().st_size > 200_000:  # ~200KB sanity threshold
                    return
            except OSError:
                pass

        if time.time() - t1 > 60:
            raise RuntimeError(f"Export finished but MP4 not created properly: {out_mp4}")
        time.sleep(1)


def export_folder_ppsm_to_mp4(
    input_dir: str,
    output_dir: str,
    use_recorded_timings: bool = True,
    default_slide_seconds: int = 5,
    vertical_resolution: int = 1080,
    fps: int = 30,
    quality: int = 85,
):
    in_path = Path(input_dir)
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()

    # DispatchEx avoids reusing an existing broken COM instance
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    ppt.Visible = True
    ppt.DisplayAlerts = 0  # disable prompts/popups

    try:
        files = sorted(in_path.glob("*.ppsm"))
        if not files:
            print("No .ppsm files found.")
            return

        for ppsm in files:
            out_mp4 = out_path / (ppsm.stem + ".mp4")

            print(f"Exporting: {ppsm.name} -> {out_mp4.name}")

            # Open (FileName, ReadOnly, Untitled, WithWindow)
            pres = ppt.Presentations.Open(str(ppsm), True, False, False)

            # CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertRes, FramesPerSecond, Quality)
            pres.CreateVideo(
                str(out_mp4),
                use_recorded_timings,
                default_slide_seconds,
                vertical_resolution,
                fps,
                quality,
            )

            wait_for_video(pres, out_mp4)

            pres.Close()

        print("Done.")

    finally:
        try:
            ppt.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    base_dir = Path(__file__).resolve().parent

    export_folder_ppsm_to_mp4(
        input_dir=base_dir,
        output_dir=base_dir / "Videos",
        use_recorded_timings=True,
        default_slide_seconds=5,
        vertical_resolution=1080,
        fps=30,
        quality=85,
    )
