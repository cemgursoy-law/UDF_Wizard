#!/bin/bash
# UDF Wizard - Mac / Linux baslatici
cd "$(dirname "$0")"
APPDIR="$(pwd)"

if ! command -v python3 &>/dev/null; then
    echo "HATA: Python3 bulunamadi."
    echo "Mac icin: https://www.python.org/downloads/"
    echo "Linux icin: sudo apt install python3 python3-pip"
    read -p "Cikis icin Enter'a basin..."
    exit 1
fi

python3 -m pip install customtkinter tkinterdnd2 python-docx pdfplumber pymupdf reportlab pillow --quiet

# Masaustu klasorunu bul (Turkce/Ingilizce)
DESKTOP="$HOME/Desktop"
[ ! -d "$DESKTOP" ] && DESKTOP="$HOME/Masaüstü"
[ ! -d "$DESKTOP" ] && DESKTOP="$HOME"

if [[ "$OSTYPE" == "darwin"* ]]; then
    ICNS="$APPDIR/UDF_LOGO.icns"
    APP_BUNDLE="$DESKTOP/UDF Wizard.app"

    # PNG -> ICNS (yalnizca bir kez, macOS yerlesik araclariyla)
    if [ ! -f "$ICNS" ]; then
        ICONSET="$APPDIR/UDF_LOGO.iconset"
        mkdir -p "$ICONSET"
        sips -z 16   16   "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_16x16.png"       &>/dev/null
        sips -z 32   32   "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_16x16@2x.png"    &>/dev/null
        sips -z 32   32   "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_32x32.png"       &>/dev/null
        sips -z 64   64   "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_32x32@2x.png"    &>/dev/null
        sips -z 128  128  "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_128x128.png"     &>/dev/null
        sips -z 256  256  "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_128x128@2x.png"  &>/dev/null
        sips -z 256  256  "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_256x256.png"     &>/dev/null
        sips -z 512  512  "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_256x256@2x.png"  &>/dev/null
        sips -z 512  512  "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_512x512.png"     &>/dev/null
        sips -z 1024 1024 "$APPDIR/UDF_LOGO.png" --out "$ICONSET/icon_512x512@2x.png"  &>/dev/null
        iconutil -c icns "$ICONSET" -o "$ICNS" 2>/dev/null
        rm -rf "$ICONSET"
    fi

    # .app bundle olustur (yalnizca bir kez)
    if [ ! -d "$APP_BUNDLE" ]; then
        mkdir -p "$APP_BUNDLE/Contents/MacOS"
        mkdir -p "$APP_BUNDLE/Contents/Resources"

        # Dock / Finder ikonu
        [ -f "$ICNS" ] && cp "$ICNS" "$APP_BUNDLE/Contents/Resources/AppIcon.icns"

        # Info.plist
        cat > "$APP_BUNDLE/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>UDF Wizard</string>
    <key>CFBundleDisplayName</key>
    <string>UDF Wizard</string>
    <key>CFBundleIdentifier</key>
    <string>com.udfwizard.app</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
PLIST

        # Terminalsiz baslatici (sessizce paket gunceller, sonra uygulamayi acar)
        cat > "$APP_BUNDLE/Contents/MacOS/launcher" <<LAUNCHER
#!/bin/bash
python3 -m pip install customtkinter tkinterdnd2 python-docx pdfplumber pymupdf reportlab pillow --quiet 2>/dev/null
exec python3 "$APPDIR/udf_donusturucu.py"
LAUNCHER
        chmod +x "$APP_BUNDLE/Contents/MacOS/launcher"

        # Finder onbellegini temizle (ikon hemen gorunsun)
        touch "$APP_BUNDLE"
        killall Finder 2>/dev/null || true
    fi

else
    # Linux: .desktop dosyasi
    SHORTCUT="$DESKTOP/UDF_Wizard.desktop"
    if [ ! -f "$SHORTCUT" ]; then
        cat > "$SHORTCUT" <<EOF
[Desktop Entry]
Type=Application
Name=UDF Wizard
Exec=python3 "$APPDIR/udf_donusturucu.py"
Path=$APPDIR
Icon=$APPDIR/UDF_LOGO.png
Terminal=false
EOF
        chmod +x "$SHORTCUT"
    fi
fi

python3 "$APPDIR/udf_donusturucu.py"
