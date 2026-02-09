#!/bin/bash
# Office Pro Android Build Script
# Automates the APK build process

echo "=========================================="
echo "  Office Pro - Android Build Script"
echo "=========================================="
echo ""

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check if running on Linux or macOS
if [[ "$OSTYPE" != "linux-gnu"* ]] && [[ "$OSTYPE" != "darwin"* ]]; then
    echo -e "${RED}Error: This script must be run on Linux or macOS${NC}"
    echo "Windows users should use WSL2 or Docker"
    exit 1
fi

# Check Python
echo "Checking Python installation..."
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}Error: Python 3 is not installed${NC}"
    exit 1
fi

PYTHON_VERSION=$(python3 --version | cut -d' ' -f2)
echo -e "${GREEN}✓ Python version: $PYTHON_VERSION${NC}"

# Check Java
echo ""
echo "Checking Java installation..."
if ! command -v java &> /dev/null; then
    echo -e "${RED}Error: Java is not installed${NC}"
    echo "Please install OpenJDK 11:"
    echo "  Ubuntu: sudo apt install openjdk-11-jdk"
    echo "  macOS: brew install openjdk@11"
    exit 1
fi

JAVA_VERSION=$(java -version 2>&1 | awk -F '"' '/version/ {print $2}')
echo -e "${GREEN}✓ Java version: $JAVA_VERSION${NC}"

# Get script directory
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if we're in the right directory
if [ ! -f "buildozer.spec" ]; then
    echo -e "${RED}Error: buildozer.spec not found${NC}"
    echo "Please run this script from the Office Pro directory"
    exit 1
fi

echo -e "${GREEN}✓ Found buildozer.spec${NC}"

# Create virtual environment if it doesn't exist
if [ ! -d "buildozer_venv" ]; then
    echo ""
    echo "Creating virtual environment..."
    python3 -m venv buildozer_venv
fi

# Activate virtual environment
echo ""
echo "Activating virtual environment..."
source buildozer_venv/bin/activate

# Install buildozer and dependencies
echo ""
echo "Installing buildozer and dependencies..."
pip install --upgrade pip
pip install buildozer cython kivy

# Check disk space
echo ""
echo "Checking disk space..."
AVAILABLE_SPACE=$(df . | awk 'NR==2 {print $4}')
AVAILABLE_GB=$((AVAILABLE_SPACE / 1024 / 1024))

if [ $AVAILABLE_GB -lt 30 ]; then
    echo -e "${YELLOW}⚠ Warning: Only ${AVAILABLE_GB}GB free space available${NC}"
    echo "Buildozer requires at least 30GB free space"
    read -p "Continue anyway? (y/N): " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
else
    echo -e "${GREEN}✓ Available space: ${AVAILABLE_GB}GB${NC}"
fi

# Menu
echo ""
echo "=========================================="
echo "  Build Options"
echo "=========================================="
echo "1. Debug Build (APK)"
echo "2. Release Build (APK)"
echo "3. Build for specific architecture"
echo "4. Clean build"
echo "5. Deploy to connected device"
echo "6. View logs"
echo "7. Full workflow (build + deploy + logs)"
echo "8. Exit"
echo ""
read -p "Select option (1-8): " choice

case $choice in
    1)
        echo ""
        echo "Starting debug build..."
        echo -e "${YELLOW}This will take 30-60 minutes on first run${NC}"
        buildozer android debug
        
        if [ $? -eq 0 ]; then
            echo ""
            echo -e "${GREEN}✓ Build successful!${NC}"
            echo "APK location: bin/"
            ls -lh bin/*.apk 2>/dev/null || echo "Check bin/ directory"
        else
            echo ""
            echo -e "${RED}✗ Build failed${NC}"
            echo "Check the error messages above"
        fi
        ;;
        
    2)
        echo ""
        echo "Release build selected"
        
        # Check for keystore
        if [ ! -f "officepro.keystore" ]; then
            echo ""
            echo "Creating keystore..."
            keytool -genkey -v -keystore officepro.keystore -alias officepro -keyalg RSA -keysize 2048 -validity 10000
        fi
        
        echo ""
        echo "Starting release build..."
        buildozer android release
        
        if [ $? -eq 0 ]; then
            echo ""
            echo -e "${GREEN}✓ Release build successful!${NC}"
            ls -lh bin/*.apk 2>/dev/null || echo "Check bin/ directory"
        else
            echo ""
            echo -e "${RED}✗ Build failed${NC}"
        fi
        ;;
        
    3)
        echo ""
        echo "Select architecture:"
        echo "1. ARM64 (arm64-v8a) - Modern devices"
        echo "2. ARM32 (armeabi-v7a) - Older devices"
        echo "3. x86_64 - Emulators"
        echo "4. All architectures"
        read -p "Select (1-4): " arch_choice
        
        case $arch_choice in
            1) ARCH="arm64-v8a" ;;
            2) ARCH="armeabi-v7a" ;;
            3) ARCH="x86_64" ;;
            4) ARCH="arm64-v8a,armeabi-v7a,x86_64" ;;
            *) ARCH="arm64-v8a" ;;
        esac
        
        echo ""
        echo "Building for architecture: $ARCH"
        buildozer android debug --arch=$ARCH
        ;;
        
    4)
        echo ""
        echo "Cleaning build..."
        buildozer android clean
        echo -e "${GREEN}✓ Build cleaned${NC}"
        ;;
        
    5)
        echo ""
        echo "Deploying to connected device..."
        
        # Check if device is connected
        if ! adb devices | grep -q "device$"; then
            echo -e "${RED}✗ No device connected${NC}"
            echo "Please connect an Android device with USB debugging enabled"
            exit 1
        fi
        
        buildozer android deploy run
        
        if [ $? -eq 0 ]; then
            echo -e "${GREEN}✓ App deployed successfully!${NC}"
        else
            echo -e "${RED}✗ Deployment failed${NC}"
        fi
        ;;
        
    6)
        echo ""
        echo "Viewing logs..."
        echo "Press Ctrl+C to exit"
        buildozer android logcat | grep python
        ;;
        
    7)
        echo ""
        echo "Starting full workflow..."
        echo -e "${YELLOW}Step 1/3: Building APK...${NC}"
        buildozer android debug
        
        if [ $? -eq 0 ]; then
            echo -e "${GREEN}✓ Build successful${NC}"
            
            # Check if device is connected
            if adb devices | grep -q "device$"; then
                echo -e "${YELLOW}Step 2/3: Deploying to device...${NC}"
                buildozer android deploy run
                
                if [ $? -eq 0 ]; then
                    echo -e "${GREEN}✓ Deployment successful${NC}"
                    echo -e "${YELLOW}Step 3/3: Showing logs...${NC}"
                    echo "Press Ctrl+C to exit"
                    buildozer android logcat | grep python
                else
                    echo -e "${RED}✗ Deployment failed${NC}"
                fi
            else
                echo -e "${YELLOW}⚠ No device connected, skipping deployment${NC}"
                echo "APK available in: bin/"
            fi
        else
            echo -e "${RED}✗ Build failed${NC}"
        fi
        ;;
        
    8)
        echo "Exiting..."
        exit 0
        ;;
        
    *)
        echo -e "${RED}Invalid option${NC}"
        exit 1
        ;;
esac

# Deactivate virtual environment
deactivate

echo ""
echo "=========================================="
echo "  Build Script Complete"
echo "=========================================="
