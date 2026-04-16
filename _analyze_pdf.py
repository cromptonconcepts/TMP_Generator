import fitz

doc = fitz.open(r'c:\Users\Sanju\OneDrive - CromptonConcepts\Application development\Apps\TMP Generator\TMP_Generator\sample\CC05855-S1-CTMP-Rev 8.pdf')
print(f'Total pages: {doc.page_count}')
for i in range(min(doc.page_count, 60)):
    page = doc.load_page(i)
    text = page.get_text('text').strip()
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    preview = ' | '.join(lines[:12])
    print(f'--- Page {i+1} ---')
    print(preview[:500])
doc.close()
