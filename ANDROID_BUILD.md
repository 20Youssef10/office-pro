"""
Office Pro Android Build Instructions
Complete guide for converting Python app to Android APK
"""

# OFFICE PRO - ANDROID APK CONVERSION GUIDE

## OVERVIEW

This guide explains how to convert Office Pro from a Python desktop application
to an Android APK using Buildozer and Kivy.

## IMPORTANT NOTE

**PyQt6 is NOT compatible with Android.** You must use Kivy framework instead.
This requires creating a mobile-optimized version of the app.

---

## PREREQUISITES

### System Requirements
- **OS**: Linux (Ubuntu 20.04+ recommended) or macOS
- **RAM**: Minimum 8GB (16GB recommended)
- **Disk Space**: 30GB+ free space
- **Architecture**: x86_64 (Intel/AMD) - ARM builds require cross-compilation

### Required Software
1. **Python 3.8+**
2. **Java JDK 11+**
3. **Android SDK & NDK** (Buildozer can download automatically)
4. **Git**

---

## STEP 1: INSTALL DEPENDENCIES

### On Ubuntu/Debian:
```bash
# Update system
sudo apt update && sudo apt upgrade -y

# Install required packages
sudo apt install -y \
    python3-pip \
    python3-venv \
    git \
    zip \
    unzip \
    openjdk-11-jdk \
    autoconf \
    libtool \
    pkg-config \
    zlib1g-dev \
    libncurses5-dev \
    libncursesw5-dev \
    libtinfo5 \
    cmake \
    libffi-dev \
    libssl-dev \
    automake
```

### On macOS:
```bash
# Install Homebrew if not already installed
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install dependencies
brew install \
    python3 \
    git \
    pkg-config \
    autoconf \
    automake \
    libtool \
    openssl
```

---

## STEP 2: SET UP PYTHON ENVIRONMENT

```bash
# Create project directory
cd "Office Pro"

# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install buildozer and cython
pip install buildozer cython

# Install Kivy and other mobile requirements
pip install kivy kivy-garden
pip install android
```

---

## STEP 3: PREPARE THE PROJECT

### File Structure for Android:
```
Office Pro/
├── buildozer.spec          # Build configuration (provided)
├── main_android.py         # Android entry point (provided)
├── modules/
│   ├── __init__.py
│   ├── word_processor_mobile.py    # Mobile-optimized
│   ├── spreadsheet_mobile.py       # Mobile-optimized
│   ├── pdf_editor_mobile.py        # Mobile-optimized
│   └── file_manager.py
├── assets/
│   ├── icon.png            # App icon (512x512)
│   ├── presplash.png       # Loading screen (optional)
│   └── fonts/              # Custom fonts (optional)
└── android/
    └── custom_rules.py     # Custom build rules (optional)
```

### Create Required Directories:
```bash
mkdir -p assets android
```

### Add App Icon:
```bash
# Download or create an icon (512x512 pixels)
# Place it in assets/icon.png
```

---

## STEP 4: CONFIGURE BUILDOZER

The `buildozer.spec` file is already provided. Key settings:

```ini
[app]
title = Office Pro
package.name = officepro
package.domain = com.officepro.app
source.dir = .
version = 1.0.0

# Requirements
requirements = python3,kivy,pyjnius,android

# Permissions
android.permissions = INTERNET,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE

# Android API
android.api = 33
android.minapi = 21
```

### Edit buildozer.spec if needed:
```bash
nano buildozer.spec
```

---

## STEP 5: BUILD THE APK

### Initial Build (Downloads SDK/NDK automatically):
```bash
# Debug build
buildozer android debug

# This will take 30-60 minutes on first run
# It downloads Android SDK, NDK, and builds all dependencies
```

### Troubleshooting First Build:
If the build fails, try:
```bash
# Clean and rebuild
buildozer android clean
buildozer android debug

# Update buildozer and dependencies
pip install --upgrade buildozer cython

# Check dependencies
buildozer android debug deploy run logcat
```

### Release Build (Signed APK):
```bash
# Generate keystore (one-time)
keytool -genkey -v -keystore officepro.keystore -alias officepro -keyalg RSA -keysize 2048 -validity 10000

# Build release APK
buildozer android release
```

---

## STEP 6: DEPLOY TO ANDROID DEVICE

### Option A: USB Connection
```bash
# Enable USB debugging on your Android device
# Connect via USB

# Deploy and run
buildozer android debug deploy run

# View logs
buildozer android logcat
```

### Option B: Manual Installation
```bash
# Find the APK
ls -la bin/

# Copy to device
adb install bin/officepro-1.0.0-arm64-v8a_armeabi-v7a-debug.apk

# Or manually copy APK to device and install
```

---

## BUILD OPTIONS

### Build for specific architectures:
```bash
# ARM 64-bit (modern devices)
buildozer android debug --arch=arm64-v8a

# ARM 32-bit (older devices)
buildozer android debug --arch=armeabi-v7a

# x86 (emulators)
buildozer android debug --arch=x86

# Build all architectures
buildozer android debug --arch=arm64-v8a,armeabi-v7a,x86_64
```

### Build Android App Bundle (AAB) for Play Store:
```bash
buildozer android release --arch=arm64-v8a,armeabi-v7a,x86_64
# Output: bin/*.aab
```

---

## COMMON ISSUES & SOLUTIONS

### Issue 1: Build fails with "No module named 'android'"
**Solution:**
```bash
pip install android
# Or add to requirements in buildozer.spec:
# requirements = python3,kivy,android
```

### Issue 2: Java version error
**Solution:**
```bash
# Check Java version
java -version

# Install OpenJDK 11
sudo apt install openjdk-11-jdk
sudo update-alternatives --config java
# Select Java 11
```

### Issue 3: Out of memory during build
**Solution:**
```bash
# Increase swap space
sudo fallocate -l 8G /swapfile
sudo chmod 600 /swapfile
sudo mkswap /swapfile
sudo swapon /swapfile

# Monitor memory
free -h
```

### Issue 4: Permission denied errors
**Solution:**
```bash
# Make scripts executable
chmod +x buildozer.spec
chmod +x main_android.py

# Run with sudo if needed (not recommended)
sudo buildozer android debug
```

### Issue 5: App crashes on startup
**Solution:**
```bash
# Check logs
buildozer android logcat | grep python

# Common fixes:
# 1. Ensure all imports work
# 2. Check for missing files in package
# 3. Verify permissions in buildozer.spec
```

### Issue 6: Module not found (e.g., docx, openpyxl)
**Solution:**
Add pure Python packages to requirements:
```ini
# In buildozer.spec
requirements = python3,kivy,python-docx,openpyxl,pandas,pillow
```

For packages with C extensions, you may need recipes:
```ini
# Check if recipe exists
buildozer recipes

# Add to requirements
requirements = python3,kivy,numpy,pandas
```

---

## ADVANCED CONFIGURATION

### Custom Recipes
If you need to compile C extensions:

1. Create `p4a-recipes/` directory
2. Add recipe for the package
3. Reference in buildozer.spec:
```ini
p4a.local_recipes = ./p4a-recipes/
```

### Reduce APK Size
```ini
# In buildozer.spec
# Exclude unnecessary files
source.exclude_patterns = 
    assets/*.psd,
    assets/*.ai,
    tests/*,
    docs/*,
    *.md,
    .git/*

# Build for specific architecture only
android.archs = arm64-v8a
```

### ProGuard (Code obfuscation)
```ini
# In buildozer.spec
android.minify_enabled = True
```

---

## PUBLISHING TO GOOGLE PLAY STORE

### Step 1: Create Signed AAB
```bash
# Generate keystore
keytool -genkey -v -keystore officepro.keystore -alias officepro -keyalg RSA -keysize 2048 -validity 10000

# Build release AAB
buildozer android release
```

### Step 2: Prepare Store Listing
1. Create Google Play Developer account ($25)
2. Create app listing
3. Upload AAB file
4. Add screenshots (phone + tablet)
5. Write description
6. Set privacy policy
7. Configure content rating

### Step 3: Testing
```bash
# Build internal testing version
buildozer android release --arch=arm64-v8a

# Upload to Play Console > Internal Testing
```

---

## ALTERNATIVE: USE DOCKER (EASIER)

If you have issues with native builds, use Docker:

```bash
# Pull buildozer image
docker pull kivy/buildozer

# Run build in container
docker run --rm -v "$(pwd)":/home/user/app kivy/buildozer android debug

# Or with current directory mounted
docker run -it --rm \
    -v "$(pwd)":/home/user/hostcwd \
    kivy/buildozer android debug
```

---

## QUICK START COMMAND CHEAT SHEET

```bash
# Setup (one-time)
cd "Office Pro"
python3 -m venv venv
source venv/bin/activate
pip install buildozer cython kivy

# Build debug APK
buildozer android debug

# Build and install on device
buildozer android debug deploy run

# View logs
buildozer android logcat

# Clean build
buildozer android clean

# Release build
buildozer android release

# Build for specific arch
buildozer android debug --arch=arm64-v8a
```

---

## LIMITATIONS ON ANDROID

### PyQt6 Not Supported
- **Problem**: PyQt6 doesn't work on Android
- **Solution**: Using Kivy framework instead
- **Impact**: Different UI, mobile-optimized interface

### File Access
- **Scoped Storage**: Android 10+ restricts file access
- **Solution**: Use app-specific directories or request permissions
- **Impact**: Files saved to `/sdcard/Android/data/com.officepro.app/files/`

### Performance
- **Python on mobile**: Slower than native apps
- **Solution**: Optimize code, use native libraries when possible
- **Impact**: Large documents may load slowly

### Background Processing
- **Android restrictions**: Apps paused in background
- **Solution**: Implement proper on_pause/on_resume handlers
- **Impact**: Auto-save works only when app is active

---

## FEATURE COMPARISON

| Feature | Desktop (PyQt6) | Mobile (Kivy) |
|---------|----------------|---------------|
| UI Framework | PyQt6 | Kivy |
| File Menu | Full menu bar | Mobile-optimized |
| Text Editing | Full rich text | Basic text editing |
| Spreadsheet | Full grid | Simplified grid |
| Charts | PyQt Charts | Kivy charts |
| PDF View | Full viewer | Basic viewer |
| Print | Full support | Not available |
| Undo/Redo | Full support | Limited support |
| Auto-save | 5 minutes | Manual + on pause |
| Image Insert | Full support | Gallery picker |

---

## TESTING CHECKLIST

Before releasing:

- [ ] App launches without crashes
- [ ] Can create new documents
- [ ] Can open existing files
- [ ] Can save files
- [ ] Text editing works
- [ ] All buttons respond
- [ ] Back button works correctly
- [ ] App handles rotation
- [ ] App works offline
- [ ] No memory leaks
- [ ] Properly requests permissions
- [ ] Handles low memory
- [ ] Works on different screen sizes

---

## SUPPORT & RESOURCES

### Documentation
- Buildozer: https://buildozer.readthedocs.io/
- Kivy: https://kivy.org/doc/stable/
- Python for Android: https://python-for-android.readthedocs.io/

### Forums
- Kivy Discord: https://chat.kivy.org/
- Reddit r/kivy: https://reddit.com/r/kivy
- Stack Overflow: Tag with [kivy] [buildozer]

### GitHub Issues
- Buildozer: https://github.com/kivy/buildozer/issues
- python-for-android: https://github.com/kivy/python-for-android/issues

---

## LICENSE

Office Pro Android version maintains the same license as the desktop version.

---

**Good luck with your Android build!** 🚀

For issues, check the logcat output:
```bash
buildozer android logcat | grep python
```
