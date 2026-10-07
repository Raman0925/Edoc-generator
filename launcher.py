"""Desktop entry point plus a packaged-app startup check for Windows CI."""

import argparse
import json
from pathlib import Path

from edoc import __version__
from edoc.gui import Application


def main():
    parser = argparse.ArgumentParser(description="eDoc Generator")
    parser.add_argument("--smoke-test-output", type=Path, help="Internal Windows build startup check")
    parser.add_argument("--smoke-screenshot-dir", type=Path, help="Optional Windows UI verification images")
    args = parser.parse_args()
    app = Application()
    if args.smoke_test_output:
        # Check that the frozen build contains Tk and the Windows COM modules.
        # This does not open Excel or claim to test Office rendering.
        import pythoncom
        import pywintypes
        import win32com.client

        def complete():
            result = {"ok": len(app.tabs.tabs()) == 3, "version": __version__,
                      "tabs": len(app.tabs.tabs()), "excel_integration_tested": False}
            if args.smoke_screenshot_dir:
                from PIL import ImageGrab
                from edoc.config import Mapping, Profile

                args.smoke_screenshot_dir.mkdir(parents=True, exist_ok=True)
                app.apply_profile(Profile(
                    template="C:/Calibration/Approval_Template.xlsx",
                    input_dir="C:/Calibration/Input", output_dir="C:/Calibration/Output",
                    approval_sheet="Approval", report_sheet="Report", mappings=[
                        Mapping("Serial number", "Report", "B4", "Approval", "F6", "text"),
                        Mapping("Calibration date", "Report", "D5", "Approval", "F7"),
                        Mapping("Sensitivity", "Report", "C12", "Approval", "D14"),
                        Mapping("Measured value", "Report", "C13", "Approval", "D15"),
                        Mapping("Uncertainty", "Report", "C14", "Approval", "D16")]))
                app.status.set("Example setup for interface verification. No report has been processed.")
                for name, panel in (("setup", app.setup_tab), ("mappings", app.mapping_tab), ("results", app.run_tab)):
                    app.tabs.select(panel)
                    app.update()
                    ImageGrab.grab(window=app.winfo_id()).save(args.smoke_screenshot_dir / f"ui-{name}.png")
            args.smoke_test_output.write_text(json.dumps(result), encoding="utf-8")
            app.destroy()

        app.after(350, complete)
    app.mainloop()


if __name__ == "__main__":
    main()
