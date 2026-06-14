#!/bin/bash
# UDF Wizard - Mac / Linux baslatici
#
# Mac'te "guvenli degil" uyarisi aliyorsaniz:
#   1. Bu dosyaya sag tiklayin -> Ac -> Ac
#   VEYA Terminal'de: xattr -d com.apple.quarantine "$(dirname "$0")/Baslat_Mac_Linux.command"
#   Sonra tekrar calistirin.

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

    # .app bundle yapi dosyalari (yalnizca yoksa olustur)
    if [ ! -d "$APP_BUNDLE" ]; then
        mkdir -p "$APP_BUNDLE/Contents/MacOS"
        mkdir -p "$APP_BUNDLE/Contents/Resources"

        [ -f "$ICNS" ] && cp "$ICNS" "$APP_BUNDLE/Contents/Resources/AppIcon.icns"

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
    <key>NSPrincipalClass</key>
    <string>NSApplication</string>
</dict>
</plist>
PLIST
    fi

    # Launcher her zaman yeniden olustur (python3 yolu degisebilir)
    PYTHON3_PATH="$(command -v python3)"
    mkdir -p "$APP_BUNDLE/Contents/MacOS"
    cat > "$APP_BUNDLE/Contents/MacOS/launcher" <<LAUNCHER
#!/bin/bash
"$PYTHON3_PATH" -m pip install customtkinter tkinterdnd2 python-docx pdfplumber pymupdf reportlab pillow --quiet 2>/dev/null
exec "$PYTHON3_PATH" "$APPDIR/udf_donusturucu.py"
LAUNCHER
    chmod +x "$APP_BUNDLE/Contents/MacOS/launcher"

    # Erisim ayricaliklari: bundle icin dogru izinleri ata
    chmod -R 755 "$APP_BUNDLE" 2>/dev/null || true

    # Gatekeeper: proje klasoru ve app bundle'dan karantina etiketini kaldir
    xattr -cr "$APPDIR" 2>/dev/null || true
    xattr -cr "$APP_BUNDLE" 2>/dev/null || true
    # Ad-hoc kod imzasi (Apple Developer hesabi gerekmez, Gatekeeper'i gecmek icin yeterli)
    codesign --deep --force --sign - "$APP_BUNDLE" 2>/dev/null || true

    # Finder ikonunu yenile
    touch "$APP_BUNDLE"
    osascript -e "tell application \"Finder\" to update item POSIX file \"$APP_BUNDLE\" as alias" 2>/dev/null || true

else
    # Linux: .desktop dosyasini her zaman yeniden olustur
    PYTHON3_PATH="$(command -v python3)"
    SHORTCUT="$DESKTOP/UDF_Wizard.desktop"
    cat > "$SHORTCUT" <<EOF
[Desktop Entry]
Type=Application
Name=UDF Wizard
Exec="$PYTHON3_PATH" "$APPDIR/udf_donusturucu.py"
Path=$APPDIR
Icon=$APPDIR/UDF_LOGO.png
Terminal=false
EOF
    chmod +x "$SHORTCUT"
fi

python3 "$APPDIR/udf_donusturucu.py"
