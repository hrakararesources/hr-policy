from docx import Document
from docx.oxml.ns import qn

doc = Document('content/work-rules/01_Work Rules and Regulations TH.docx')
count = 0
for para in doc.paragraphs:
    t = para.text.strip()
    if not t:
        continue
    numPr = para._element.find('.//' + qn('w:numPr'))
    if numPr is not None:
        ilvl = numPr.find(qn('w:ilvl'))
        numId = numPr.find(qn('w:numId'))
        level = ilvl.get(qn('w:val')) if ilvl is not None else '?'
        nid = numId.get(qn('w:val')) if numId is not None else '?'
        print(f'level={level} numId={nid} | {t[:60]}')
        count += 1
        if count >= 20:
            break