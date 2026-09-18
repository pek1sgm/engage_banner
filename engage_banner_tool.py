#!/usr/bin/env python3
"""
Viva Engage Community-Banner Tool
===================================
Wandelt ein aus PowerPoint exportiertes Folienbild in ein perfektes
Community-Cover für Microsoft Viva Engage (Teams Engage) um.

Außerdem kann es die Windows-Registry automatisch so setzen, dass
PowerPoint beim nächsten "Speichern als Bild" in hoher Auflösung
    exportiert (600 DPI statt der lausigen 96 DPI).

Zielmaße Viva Engage Community Cover:
    - Empfohlen:  2800 × 1048 px (2×-Variante für schärferen Engage-Upload)
  - Minimum:     700 × 262 px
  - Seitenverhältnis: ≈ 2.672 : 1
  - Max. Dateigröße: 20 MB
  - Formate: PNG, JPEG, nicht-animiertes GIF

═══════════════════════════════════════════════════════════════
WORKFLOW (2 Schritte):

  Schritt 1 — Einmalig: PowerPoint Export-DPI hochsetzen
  ─────────────────────────────────────────────────────────
    python engage_banner_tool.py --setup-registry

    → Erstellt eine .reg-Datei und bietet an, die Registry
    direkt zu setzen. Danach exportiert PowerPoint mit 600 DPI
    (= 8000×4500 px bei 16:9-Widescreen).

  Schritt 2 — Pro Bild: Export nachbearbeiten
  ─────────────────────────────────────────────────────────
    python engage_banner_tool.py mein_slide_export.png

    → Schneidet auf 2800×1048 zu und skaliert mit Lanczos.
      Ergebnis: engage_banner.png (oder --output Name)

═══════════════════════════════════════════════════════════════

Optionen:
  mein_bild.png       Eingabebild (PNG/JPG aus PowerPoint-Export)
  --output PATH       Ausgabepfad (Standard: engage_banner.png)
    --width W           Zielbreite in px (Standard: 2800)
    --height H          Zielhöhe in px (Standard: 1048)
  --format FMT        png oder jpg (Standard: png)
  --quality Q         JPEG-Qualität 1-100 (Standard: 95)
    --setup-registry    PowerPoint Export-DPI auf 600 setzen
  --preview           Zeigt Vorschau der Crop-Zone vor dem Speichern
  --gravity POS       Vertikaler Anker: center(=Oberkante)/top/bottom (Standard: center)

Voraussetzungen:
  pip install Pillow
"""

import argparse
import os
import sys
import platform

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("❌ Pillow nicht installiert!")
    print("   Bitte ausführen: pip install Pillow")
    sys.exit(1)


# ──────────────────────────────────────────────────────────────
# Konfiguration
# ──────────────────────────────────────────────────────────────
DEFAULT_WIDTH = 2800
DEFAULT_HEIGHT = 1048
DEFAULT_FORMAT = "png"
DEFAULT_QUALITY = 95


# ──────────────────────────────────────────────────────────────
# Registry-Setup: PowerPoint Export-DPI hochsetzen
# ──────────────────────────────────────────────────────────────
REG_FILE_CONTENT = r"""Windows Registry Editor Version 5.00

; PowerPoint Export-Auflösung auf 600 DPI setzen
; (Standard ist 96 DPI = nur 1280x720 bei Widescreen)
; Bei 600 DPI: 8000x4500 px — maximale Ausgangsauflösung
;
; Gilt für: PowerPoint 2016, 2019, 2021, Microsoft 365
; Pfad: HKCU\Software\Microsoft\Office\16.0\PowerPoint\Options

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options]
"ExportBitmapResolution"=dword:00000258
"""
# 0x258 = 600 dezimal


def setup_registry():
    """Erstellt .reg-Datei und setzt optional die Registry direkt."""
    if platform.system() != "Windows":
        print("⚠️  Registry-Setup ist nur unter Windows nötig.")
        print("   Auf macOS: PowerPoint > Einstellungen > Foliegröße")
        print("   anpassen und dann als PNG exportieren.\n")
        # Trotzdem die .reg-Datei erstellen, falls User sie übertragen will
        reg_path = "powerpoint_600dpi.reg"
        with open(reg_path, "w", encoding="utf-8") as f:
            f.write(REG_FILE_CONTENT)
        print(f"📄 .reg-Datei erstellt: {reg_path}")
        print("   (Auf einen Windows-PC übertragen und doppelklicken)")
        return

    reg_path = "powerpoint_600dpi.reg"
    with open(reg_path, "w", encoding="utf-8") as f:
        f.write(REG_FILE_CONTENT)
    print(f"📄 .reg-Datei erstellt: {reg_path}\n")

    print("Möchtest du die Registry jetzt direkt setzen?")
    print("(PowerPoint muss danach neu gestartet werden)\n")

    try:
        choice = input("Registry jetzt setzen? [j/N]: ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        choice = "n"

    if choice in ("j", "ja", "y", "yes"):
        import subprocess
        try:
            result = subprocess.run(
                [
                    "reg", "add",
                    r"HKCU\Software\Microsoft\Office\16.0\PowerPoint\Options",
                    "/v", "ExportBitmapResolution",
                    "/t", "REG_DWORD",
                    "/d", "600",
                    "/f",
                ],
                capture_output=True, text=True,
            )
            if result.returncode == 0:
                print("\n✅ Registry gesetzt! ExportBitmapResolution = 600")
                print("   → PowerPoint exportiert jetzt mit 600 DPI")
                print("   → Widescreen-Folien: 8000 × 4500 px")
                print("\n⚠️  PowerPoint jetzt schließen und neu öffnen!")
            else:
                print(f"\n❌ Fehler: {result.stderr}")
                print(f"   Alternativ: Doppelklick auf {reg_path}")
        except FileNotFoundError:
            print("\n❌ 'reg' Befehl nicht gefunden.")
            print(f"   Alternativ: Doppelklick auf {reg_path}")
    else:
        print(f"\n👉 Kein Problem! Doppelklicke später auf: {reg_path}")

    print("\n" + "─" * 55)
    print("NÄCHSTER SCHRITT:")
    print("  1. PowerPoint öffnen (neu starten falls offen)")
    print("  2. Datei → Speichern unter → PNG auswählen")
    print("  3. 'Nur diese Folie' wählen")
    print("  4. Dann dieses Tool mit dem exportierten Bild aufrufen:")
    print("     python engage_banner_tool.py mein_export.png")
    print("─" * 55)


# ──────────────────────────────────────────────────────────────
# Bild-Verarbeitung: Zuschnitt + Skalierung
# ──────────────────────────────────────────────────────────────
def analyze_image(img_path: str) -> Image.Image:
    """Öffnet das Bild und zeigt eine Analyse."""
    img = Image.open(img_path)
    w, h = img.size
    ratio = w / h
    dpi_info = img.info.get("dpi", (96, 96))

    print(f"\n📊 Bildanalyse: {os.path.basename(img_path)}")
    print(f"   Größe:            {w} × {h} px")
    print(f"   Seitenverhältnis: {ratio:.3f}:1")
    print(f"   DPI (Metadaten):  {dpi_info[0]:.0f}")
    print(f"   Dateigröße:       {os.path.getsize(img_path)/1024:.0f} KB")

    # Qualitätsbewertung
    if w >= 2667 and h >= 1500:
        print(f"   Qualität:         🟢 Sehr gut (≥300 DPI Export)")
    elif w >= 2000 and h >= 1125:
        print(f"   Qualität:         🟡 Gut (≥200 DPI Export)")
    elif w >= 1333 and h >= 750:
        print(f"   Qualität:         🟠 Ausreichend (≥100 DPI Export)")
    else:
        print(f"   Qualität:         🔴 Niedrig (96 DPI Standard-Export)")
        print(f"   ⚠️  Empfehlung: Erst --setup-registry ausführen,")
        print(f"      dann erneut aus PowerPoint exportieren!")

    return img


def crop_and_resize(
    img: Image.Image,
    target_width: int,
    target_height: int,
    output_path: str,
    fmt: str = "png",
    quality: int = 95,
    gravity: str = "center",
    preview: bool = False,
) -> str:
    """
    Schneidet auf Viva-Engage-Seitenverhältnis zu und skaliert
    auf die exakte Zielgröße mit Lanczos-Resampling.
    """
    src_w, src_h = img.size
    target_ratio = target_width / target_height  # ≈ 2.672
    src_ratio = src_w / src_h

    print(f"\n✂️  Zuschnitt für {target_width}×{target_height} ...")

    if abs(src_ratio - target_ratio) < 0.01:
        # Seitenverhältnis passt schon
        cropped = img
        print("   Seitenverhältnis passt — kein Zuschnitt nötig")
    elif src_ratio < target_ratio:
        # Folie ist höher als das Ziel → unten abschneiden (Oberkante halten)
        new_h = int(src_w / target_ratio)
        if gravity == "top":
            offset_y = 0
        elif gravity == "bottom":
            offset_y = src_h - new_h
        else:  # center — vertikal trotzdem Oberkante halten
            offset_y = 0

        crop_box = (0, offset_y, src_w, offset_y + new_h)
        cropped = img.crop(crop_box)
        removed = src_h - new_h
        print(f"   Höhe beschnitten: {src_h} → {new_h} px "
              f"(−{removed} px von unten, Oberkante gehalten)")

        if preview:
            _show_preview(img, crop_box)

    else:
        # Folie ist breiter als das Ziel → links/rechts beschneiden
        new_w = int(src_h * target_ratio)
        offset_x = (src_w - new_w) // 2
        crop_box = (offset_x, 0, offset_x + new_w, src_h)
        cropped = img.crop(crop_box)
        removed = src_w - new_w
        print(f"   Breite beschnitten: {src_w} → {new_w} px (−{removed} px)")

        if preview:
            _show_preview(img, crop_box)

    # Skalierung mit Lanczos (höchste Qualität)
    final = cropped.resize((target_width, target_height), Image.LANCZOS)
    print(f"   Skaliert auf {target_width}×{target_height} px (Lanczos)")

    # Speichern
    save_kwargs = {"dpi": (300, 300)}
    if fmt.lower() == "png":
        save_kwargs["optimize"] = True
    elif fmt.lower() in ("jpg", "jpeg"):
        fmt = "jpeg"
        save_kwargs["quality"] = quality
        save_kwargs["optimize"] = True
        if final.mode == "RGBA":
            background = Image.new("RGB", final.size, (255, 255, 255))
            background.paste(final, mask=final.split()[3])
            final = background

    final.save(output_path, format=fmt.upper(), **save_kwargs)

    file_size_kb = os.path.getsize(output_path) / 1024
    print(f"\n{'═'*55}")
    print(f"  ✅ FERTIG — Banner bereit für Viva Engage!")
    print(f"{'═'*55}")
    print(f"  📁 Datei:    {output_path}")
    print(f"  📐 Größe:    {target_width} × {target_height} px")
    print(f"  💾 Datei:    {file_size_kb:.0f} KB")
    print(f"  📋 Format:   {fmt.upper()}")

    if file_size_kb > 20 * 1024:
        print(f"\n  ⚠️  {file_size_kb/1024:.1f} MB — Viva Engage max. 20 MB!")
        print("     Tipp: --format jpg --quality 85")

    print(f"\n  Upload: Community → Header → 'Upload Cover Photo'")

    return output_path


def _show_preview(img: Image.Image, crop_box):
    """Zeichnet die Crop-Zone ins Originalbild und zeigt es."""
    preview = img.copy()
    draw = ImageDraw.Draw(preview, "RGBA")

    # Abgedunkelter Bereich außerhalb der Crop-Zone
    w, h = img.size
    x1, y1, x2, y2 = crop_box

    # Oben
    if y1 > 0:
        draw.rectangle([0, 0, w, y1], fill=(0, 0, 0, 128))
    # Unten
    if y2 < h:
        draw.rectangle([0, y2, w, h], fill=(0, 0, 0, 128))
    # Links
    if x1 > 0:
        draw.rectangle([0, y1, x1, y2], fill=(0, 0, 0, 128))
    # Rechts
    if x2 < w:
        draw.rectangle([x2, y1, w, y2], fill=(0, 0, 0, 128))

    # Crop-Rahmen
    draw.rectangle([x1, y1, x2, y2], outline=(0, 200, 0, 255), width=4)

    preview_path = "engage_crop_preview.png"
    preview.save(preview_path)
    print(f"\n   👁️  Vorschau gespeichert: {preview_path}")
    print("      (Grüner Rahmen = sichtbarer Bereich im Banner)")


# ──────────────────────────────────────────────────────────────
# Hauptprogramm
# ──────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Viva Engage Community-Banner aus PowerPoint-Export erstellen",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "image", nargs="?",
        help="Bild aus PowerPoint-Export (PNG/JPG)",
    )
    parser.add_argument("--output", default="engage_banner.png",
                        help="Ausgabepfad (Standard: engage_banner.png)")
    parser.add_argument("--width", type=int, default=DEFAULT_WIDTH,
                        help=f"Zielbreite px (Standard: {DEFAULT_WIDTH})")
    parser.add_argument("--height", type=int, default=DEFAULT_HEIGHT,
                        help=f"Zielhöhe px (Standard: {DEFAULT_HEIGHT})")
    parser.add_argument("--format", choices=["png", "jpg"], default=DEFAULT_FORMAT,
                        help="Ausgabeformat (Standard: png)")
    parser.add_argument("--quality", type=int, default=DEFAULT_QUALITY,
                        help="JPEG-Qualität 1-100 (Standard: 95)")
    parser.add_argument("--gravity", choices=["center", "top", "bottom"],
                        default="center",
                        help="Crop-Schwerpunkt (Standard: center)")
    parser.add_argument("--preview", action="store_true",
                        help="Crop-Vorschau speichern")
    parser.add_argument("--setup-registry", action="store_true",
                        help="PowerPoint Export-DPI auf 600 setzen (einmalig)")

    args = parser.parse_args()

    print("╔═══════════════════════════════════════════════════╗")
    print("║   Viva Engage Community-Banner Tool               ║")
    print("║   Ziel: 2800×1048 px (2×, schärferer Upload)      ║")
    print("╚═══════════════════════════════════════════════════╝")

    if args.setup_registry:
        setup_registry()
        return

    if not args.image:
        parser.print_help()
        print("\n" + "─" * 55)
        print("SCHNELLSTART:")
        print("  1. python engage_banner_tool.py --setup-registry")
        print("  2. PowerPoint: Datei → Speichern unter → PNG")
        print("  3. python engage_banner_tool.py Folie1.png")
        print("─" * 55)
        return

    if not os.path.exists(args.image):
        print(f"\n❌ Datei nicht gefunden: {args.image}")
        sys.exit(1)

    # Analyse
    img = analyze_image(args.image)

    # Zuschnitt + Skalierung
    crop_and_resize(
        img,
        args.width,
        args.height,
        args.output,
        fmt=args.format,
        quality=args.quality,
        gravity=args.gravity,
        preview=args.preview,
    )


if __name__ == "__main__":
    main()