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

python3 -m pip install customtkinter python-docx pdfplumber reportlab pillow --quiet

# Masaustu klasorunu bul (Turkce/Ingilizce)
DESKTOP="$HOME/Desktop"
[ ! -d "$DESKTOP" ] && DESKTOP="$HOME/Masaüstü"
[ ! -d "$DESKTOP" ] && DESKTOP="$HOME"

if [[ "$OSTYPE" == "darwin"* ]]; then
    SHORTCUT="$DESKTOP/UDF Wizard.command"
    if [ ! -e "$SHORTCUT" ]; then
        ln -s "$APPDIR/Baslat_Mac_Linux.command" "$SHORTCUT"
        chmod +x "$SHORTCUT"
    fi
else
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
