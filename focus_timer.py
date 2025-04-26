#!/usr/bin/env python3
# focus_timer.py - Launcher for the Focus Timer application

import sys
import os
from pathlib import Path

# Add the parent directory to the path so we can import the focus_timer_app package
sys.path.append(str(Path(__file__).parent))

# Import the main function from the focus_timer_app package
from focus_timer_app.main import main

if __name__ == "__main__":
    sys.exit(main())
