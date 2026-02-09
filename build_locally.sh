#!/bin/bash
# Office Pro - Quick Build Guide
# Run this on your local Linux/macOS machine with sufficient disk space

echo "=========================================="
echo "  Office Pro APK Build - Local Machine"
echo "=========================================="
echo ""
echo "⚠️  IMPORTANT: You need at least 30GB free disk space!"
echo ""

# Check if we have enough space
check_space() {
    AVAILABLE=$(df . | awk 'NR==2 {print $4}')
    AVAILABLE_GB=$((AVAILABLE / 1024 / 1024))
    
    if [ $AVAILABLE_GB -lt 30 ]; then
        echo "❌ ERROR: Only ${AVAILABLE_GB}GB available. Need at least 30GB."
        echo ""
        echo "Free up space or use an external drive."
        exit 1
    else
        echo "✓ Disk space check passed: ${AVAILABLE_GB}GB available"
    fi
}

# Main build process
main() {
    check_space
    
    echo ""
    echo "Step 1: Installing dependencies..."
    echo "==================================="
    
    # Detect OS
    if [[ "$OSTYPE" == "linux-gnu"* ]]; then
        echo "Detected: Linux"
        sudo apt update
        sudo apt install -y python3-pip python3-venv openjdk-11-jdk git zip unzip
    elif [[ "$OSTYPE" == "darwin"* ]]; then
        echo "Detected: macOS"
        brew install python3 openjdk@11 git
    else
        echo "Unsupported OS: $OSTYPE"
        exit 1
    fi
    
    echo ""
    echo "Step 2: Setting up Python environment..."
    echo "========================================="
    python3 -m venv buildozer_env
    source buildozer_env/bin/activate
    pip install --upgrade pip
    pip install buildozer cython kivy
    
    echo ""
    echo "Step 3: Starting build..."
    echo "========================="
    echo "⚠️  This will take 30-60 minutes on first run!"
    echo ""
    read -p "Press Enter to start building..."
    
    # Build debug APK
    buildozer android debug
    
    if [ $? -eq 0 ]; then
        echo ""
        echo "✅ BUILD SUCCESSFUL!"
        echo "===================="
        echo ""
        echo "Your APK is located at:"
        ls -lh bin/*.apk
        echo ""
        echo "To install on your Android device:"
        echo "  1. Enable USB debugging on your device"
        echo "  2. Connect via USB"
        echo "  3. Run: buildozer android deploy run"
        echo ""
        echo "Or manually install the APK from bin/ directory"
    else
        echo ""
        echo "❌ BUILD FAILED"
        echo "==============="
        echo "Check the error messages above."
        echo "Common fixes:"
        echo "  - Run: buildozer android clean"
        echo "  - Check buildozer.spec configuration"
        echo "  - Ensure all dependencies are installed"
    fi
}

# Alternative: Docker build
docker_build() {
    echo "Docker Build Option"
    echo "==================="
    echo ""
    echo "If you have Docker installed, you can use:"
    echo ""
    echo "  docker run --rm -v \"\$(pwd)\":/home/user/app \\"
    echo "    kivy/buildozer android debug"
    echo ""
    echo "This handles all dependencies automatically!"
}

# Menu
echo "Choose build method:"
echo "1. Local Build (requires 30GB space)"
echo "2. Docker Build (easiest, requires Docker)"
echo "3. Check requirements only"
read -p "Select option (1-3): " choice

case $choice in
    1)
        main
        ;;
    2)
        docker_build
        ;;
    3)
        check_space
        java -version
        python3 --version
        ;;
    *)
        echo "Invalid option"
        exit 1
        ;;
esac
