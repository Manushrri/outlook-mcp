#!/usr/bin/env python3
"""
Microsoft Outlook MCP Server
Run this to start the MCP server for Outlook integration.
"""

import sys
from pathlib import Path

# Ensure project root is in path
project_root = Path(__file__).resolve().parent
sys.path.insert(0, str(project_root))

from src.main import main

if __name__ == "__main__":
    main()
