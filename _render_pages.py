import fitz
import os

PDF_PATH = r'c:\Users\Sanju\OneDrive - CromptonConcepts\Application development\Apps\TMP Generator\TMP_Generator\sample\CC05855-S1-CTMP-Rev 8.pdf'
OUT_DIR = r'c:\Users\Sanju\OneDrive - CromptonConcepts\Application development\Apps\TMP Generator\TMP_Generator\sample'

doc = fitz.open(PDF_PATH)
# Render key pages to PNG for visual inspection
for page_idx in [0, 5, 10, 15, 22, 27, 43]:
    if page_idx < doc.page_count:
        page = doc.load_page(page_idx)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
        out_path = os.path.join(OUT_DIR, f'page_{page_idx + 1}.png')
        pix.save(out_path)
        print(f'Saved {out_path}')
doc.close()
