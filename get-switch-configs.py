"""
===============================================================================
Script Name: get-switch-configs.py
Author: Frank Abraham
Version: 1.1
License: MIT License

Description:
    This script connects to multiple Cisco switches using Netmiko, retrieves
    the running configuration from each device, and saves them as individual
    text files for later formatting and documentation.

    Credentials are securely stored in an INI file (creds.ini), and target
    switch IP addresses are read from a text file (switch-targets.txt).

    This script represents Phase 1 of a two-part process:
        1. get-switch-configs.py      - Gathers configurations from all switches
        2. Build-Final-Configs-Doc.ps1 - Formats and compiles all configs into
                                         a standardized Word document

Usage:
    - Place this script, creds.ini, and switch-targets.txt in the same folder.
    - Run the script from PowerShell or CMD:
        python get-switch-configs.py

Dependencies:
    pip install netmiko
===============================================================================
"""

import os
import sys
import time
import logging
from pathlib import Path
import configparser
from netmiko import ConnectHandler, NetmikoTimeoutException, NetmikoAuthenticationException

# ---------------------------------------------------------------------------
# Configuration Paths (update these paths for your local environment)
# ---------------------------------------------------------------------------
BASE = Path(r"C:\Path\To\Your\Scripts")
TARGETS = BASE / "switch-targets.txt"
CREDS = BASE / "creds.ini"
OUTDIR = BASE / "final-configs"
SSH_LOG =_
