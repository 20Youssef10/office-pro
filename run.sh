#!/bin/bash
# Office Pro Launcher Script

echo "=========================================="
echo "  Office Pro - Professional Office Suite"
echo "=========================================="
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed!"
    echo "Please install Python 3.8 or higher."
    exit 1
fi

# Get script directory
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if virtual environment exists, if not create it
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Install/Update dependencies
echo "Checking dependencies..."
pip install -q -r requirements.txt

# Run Office Pro
echo "Starting Office Pro..."
echo ""
python3 office_pro.py

# Deactivate virtual environment
deactivate
