#!/usr/bin/env python3
"""
PDF_Compress_no_gs.py

Compress a PDF WITHOUT Ghostscript by rasterizing pages to JPEG and
rebuilding a new PDF from those JPEGs.

Immediate file-selection dialogs are used (no initial messagebox).
Target size defaults to 5 MB. The script will try progressively lower
quality/scale combinations until target is met or the minimum settings
are reached, and will save the best result it produces.

Requirements:
    pip install pymupdf pillow
"""
import io
import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import fitz  # pymupdf
except Exception as e:
    raise SystemExit("pymupdf is required. Install with: pip install pymupdf") from e

try:
    from PIL import Image
except Exception as e:
    raise SystemExit("Pillow is required. Install with: pip install pillow") from e

TARGET_BYTES = 5 * 1024 * 1024  # 5 MB target
# Adjustable parameters: (scale, jpeg_quality)
# Script will try combos in order until target reached.
TRY_COMBOS = [
    (1.0, 85),
    (1.0, 75),
    (0.9, 75),
    (0.9, 65),
    (0.8, 65),
    (0.8, 55),
    (0.7, 55),
    (0.6, 50),
    (0.5, 45),
]
# Minimum safety limits
MIN_QUALITY = 30
MIN_SCALE = 0.3

def human_size(n):
    if n is None:
        return "N/A"
    for unit in ('B','KB','MB','GB'):
        if n < 1024.0:
            return f"{n:.2f}{unit}"
        n /= 1024.0
    return f"{n:.2f}TB"

def pdf_from_images_bytes(image_bytes_list):
    """
    Create a new PDF (in-memory bytes) from list of JPEG bytes using PyMuPDF.
    Returns bytes of the new PDF.
    """
    new_doc = fitz.open()
    for img_bytes in image_bytes_list:
        img = fitz.open("jpeg", img_bytes)  # open image as document
        rect = img[0].rect
        new_page = new_doc.new_page(width=rect.width, height=rect.height)
        # Insert the image so it fills the whole page
        new_page.insert_image(rect, stream=img_bytes)
        img.close()
    out_bytes = new_doc.write()
    new_doc.close()
    return out_bytes

def render_page_to_jpeg_bytes(page, scale=1.0, quality=75):
    """
    Render a fitz.Page to JPEG bytes at given scale and JPEG quality.
    Returns JPEG bytes.
    """
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, alpha=False)  # RGB
    # Convert pix to PIL Image
    mode = "RGB"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    bio = io.BytesIO()
    # save as JPEG with quality and optimization
    img.save(bio, format="JPEG", quality=int(quality), optimize=True)
    return bio.getvalue()

def try_compress(input_pdf_path, scale, quality):
    """
    Render every page to JPEG with given scale & quality and return resulting PDF bytes and total size.
    """
    doc = fitz.open(input_pdf_path)
    images = []
    for p in doc:
        try:
            jb = render_page_to_jpeg_bytes(p, scale=scale, quality=quality)
        except Exception as e:
            doc.close()
            raise
        images.append(jb)
    doc.close()
    pdf_bytes = pdf_from_images_bytes(images)
    return pdf_bytes, len(pdf_bytes)

def main():
    root = tk.Tk()
    root.withdraw()

    # Ask user to pick input PDF directly (no prior messagebox).
    input_path = filedialog.askopenfilename(title="Select PDF to compress",
                                            filetypes=[("PDF files", "*.pdf")])
    if not input_path:
        print("No input selected. Exiting.")
        return

    suggested_name = os.path.splitext(os.path.basename(input_path))[0] + "_compressed.pdf"
    output_path = filedialog.asksaveasfilename(title="Save compressed PDF as",
                                               initialfile=suggested_name,
                                               defaultextension=".pdf",
                                               filetypes=[("PDF files", "*.pdf")])
    if not output_path:
        print("No output selected. Exiting.")
        return

    try:
        orig_size = os.path.getsize(input_path)
    except Exception as e:
        messagebox.showerror("File error", f"Cannot read input file: {e}")
        return

    print(f"Input: {input_path} ({human_size(orig_size)})")
    if orig_size <= TARGET_BYTES:
        # already small
        import shutil
        shutil.copy2(input_path, output_path)
        messagebox.showinfo("Done", f"File already ≤ {human_size(TARGET_BYTES)}. Copied to:\n{output_path}")
        return

    best_bytes = None
    best_size = None
    best_combo = None

    # Iterate through combo settings
    for scale, quality in TRY_COMBOS:
        if quality < MIN_QUALITY or scale < MIN_SCALE:
            continue
        print(f"Trying scale={scale:.2f}, quality={quality} ...")
        try:
            pdf_bytes, size = try_compress(input_path, scale, quality)
        except Exception as e:
            print(f"Failed at scale={scale},quality={quality}: {e}")
            continue

        print(f"Result size: {human_size(size)}")
        # Track best
        if best_size is None or size < best_size:
            best_size = size
            best_bytes = pdf_bytes
            best_combo = (scale, quality)

        # If meets target, write and return
        if size <= TARGET_BYTES:
            with open(output_path, "wb") as f:
                f.write(pdf_bytes)
            messagebox.showinfo("Done", f"Success: scale={scale}, quality={quality}\nSaved: {output_path}\nSize: {human_size(size)}")
            print("Done.")
            return

    # None met target: save best we got (if any)
    if best_bytes:
        with open(output_path, "wb") as f:
            f.write(best_bytes)
        messagebox.showwarning("Partial result",
                               ("Could not reach target size of {:.2f}MB.\n"
                                "Best result: scale={:.2f}, quality={}, size={}.\nSaved to:\n{}")
                               .format(TARGET_BYTES/1024/1024, best_combo[0], best_combo[1], human_size(best_size), output_path))
    else:
        messagebox.showerror("Failed", "Compression failed (no valid result).")

if __name__ == "__main__":
    main()
