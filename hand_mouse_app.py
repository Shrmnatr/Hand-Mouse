#!/usr/bin/env python3
"""
hand_mouse_app.py - Launcher used to build HandMouse.exe.

Extra settings can be added after the exe name, for example:
  HandMouse.exe --far
  HandMouse.exe --monitor 2 --mouse-hand left
"""
import os
import sys
import traceback


def resource_path(name):
    """Find a bundled file, both inside the exe and when run as a script."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def run():
    print("Starting Hand Mouse, please wait...\n")
    import hand_mouse_pro as app
    app.MODEL_PATH = resource_path("hand_landmarker.task")
    sys.argv = ["HandMouse"] + sys.argv[1:]
    app.main()


if __name__ == "__main__":
    try:
        run()
    except SystemExit as e:
        if e.code not in (None, 0):
            print(f"\n{e.code}")
    except KeyboardInterrupt:
        pass
    except Exception:
        print("\nSomething went wrong:\n")
        traceback.print_exc()
        print("\nTake a screenshot of this window and send it to the person "
              "who gave you this program.")
    try:
        input("\nHand Mouse has closed. Press Enter to close this window...")
    except (EOFError, KeyboardInterrupt):
        pass
