"""
Genera un file icon.ico per l'applicazione Mail Merge.
Richiede: pip install Pillow
Esegui: python crea_icona.py
"""

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow", "-q"])
    from PIL import Image, ImageDraw, ImageFont

import os

def crea_icona():
    sizes = [256, 128, 64, 48, 32, 16]
    frames = []

    for size in sizes:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Sfondo arrotondato indigo
        r = size // 6
        draw.rounded_rectangle([0, 0, size-1, size-1], radius=r,
                                fill=(79, 70, 229, 255))

        # Documento (rettangolo bianco con angolo piegato)
        m  = size // 5
        dw = size - m * 2
        dh = int(dw * 1.3)
        dy = (size - dh) // 2
        fold = size // 6

        # Corpo documento
        draw.rectangle([m, dy, m + dw, dy + dh], fill=(255, 255, 255, 240))

        # Angolo piegato (triangolo)
        draw.polygon([
            (m + dw - fold, dy),
            (m + dw,        dy),
            (m + dw,        dy + fold)
        ], fill=(200, 196, 240, 255))

        # Linee testo
        lm   = m + size // 8
        lw   = dw - size // 4
        lh   = max(1, size // 16)
        lgap = size // 10
        ly   = dy + size // 6

        for i in range(4):
            w = lw if i < 3 else lw * 2 // 3
            draw.rectangle([lm, ly, lm + w, ly + lh],
                           fill=(79, 70, 229, 200))
            ly += lgap

        frames.append(img)

    frames[0].save(
        "icon.ico",
        format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=frames[1:]
    )
    print("icon.ico creato con successo!")

if __name__ == "__main__":
    crea_icona()