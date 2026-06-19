#!/usr/bin/env python3
"""
Akara Resources — HR Policy HTML Generator v3
แบ่งเนื้อหาตามหน้า PDF (page-based splitting) แทน text pattern
"""

import os, re

try:
    import fitz
    PDF_LIB = "pymupdf"
except ImportError:
    PDF_LIB = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# CONFIG — กำหนดว่าแต่ละหมวดอยู่หน้าไหนของ PDF
# ดูจากสารบัญใน PDF: หมวดที่ 1 = หน้า 4, หมวดที่ 2 = หน้า 8 ...
# (หน้า PDF เริ่มจาก 0 แต่เราใช้เลขจริงจากสารบัญ -1)
# ============================================================
WORK_RULES_CHAPTERS = [
    {"num": 0,  "title_th": "สารจากผู้บริหาร",
     "title_en": "Message from Management",
     "page_start": 2, "page_end": 2},
    {"num": 1,  "title_th": "บททั่วไป",
     "title_en": "General Provisions",
     "page_start": 3, "page_end": 6},
    {"num": 2,  "title_th": "วิสัยทัศน์องค์กร เสาหลักกลยุทธ์และนโยบายองค์กร",
     "title_en": "Corporate Vision, Strategic Pillars and Policies",
     "page_start": 7, "page_end": 9},
    {"num": 3,  "title_th": "การว่าจ้าง",
     "title_en": "Employment",
     "page_start": 10, "page_end": 11},
    {"num": 4,  "title_th": "การสับเปลี่ยนหน้าที่การปฏิบัติงาน การโยกย้าย และการปฏิบัติงานที่บ้าน",
     "title_en": "Job Rotation, Transfer and Work from Home",
     "page_start": 12, "page_end": 13},
    {"num": 5,  "title_th": "วันทำงาน เวลาทำงานปกติ และเวลาพัก",
     "title_en": "Working Days, Hours and Rest Periods",
     "page_start": 14, "page_end": 16},
    {"num": 6,  "title_th": "วันหยุด และหลักเกณฑ์วันหยุด",
     "title_en": "Holidays and Holiday Criteria",
     "page_start": 17, "page_end": 19},
    {"num": 7,  "title_th": "หลักเกณฑ์การทำงานล่วงเวลา และการทำงานในวันหยุด",
     "title_en": "Overtime and Holiday Work Criteria",
     "page_start": 20, "page_end": 21},
    {"num": 8,  "title_th": "การจ่ายค่าจ้าง ค่าล่วงเวลา และค่าทำงานในวันหยุด",
     "title_en": "Wages, Overtime Pay and Holiday Pay",
     "page_start": 22, "page_end": 26},
    {"num": 9,  "title_th": "วันลา และหลักเกณฑ์การลา",
     "title_en": "Leave and Leave Criteria",
     "page_start": 27, "page_end": 33},
    {"num": 10, "title_th": "วินัย และโทษทางวินัย",
     "title_en": "Discipline and Disciplinary Penalties",
     "page_start": 34, "page_end": 44},
    {"num": 11, "title_th": "การร้องทุกข์",
     "title_en": "Grievance Procedure",
     "page_start": 45, "page_end": 46},
    {"num": 12, "title_th": "การพ้นสภาพการเป็นพนักงาน",
     "title_en": "Termination of Employment",
     "page_start": 47, "page_end": 51},
    {"num": 13, "title_th": "สวัสดิการและสิทธิประโยชน์อื่น ๆ",
     "title_en": "Welfare and Other Benefits",
     "page_start": 52, "page_end": 53},
    {"num": 14, "title_th": "ข้อกำหนดทั่วไปและการบังคับใช้",
     "title_en": "General Provisions and Enforcement",
     "page_start": 54, "page_end": 55},
]

JUNK = [
    r'ห้ามทำสำเนา[^\n]*', r'ห้ามท า[^\n]*',
    r'AKR-DCC[^\n]*', r'Support Document[^\n]*',
    r'Document No\.[^\n]*', r'Document Title:[^\n]*',
    r'Revision No\.[^\n]*', r'Effective Date[^\n]*',
    r'Page No\.[^\n]*', r'Work Rules and Regulations\s*',
    r'Welfare and Benefits\s*',
]

def clean(text):
    for p in JUNK:
        text = re.sub(p, '', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def extract_pages(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []
    for page in doc:
        pages.append(clean(page.get_text("text")))
    doc.close()
    return pages

def get_chapter_text(pages, ch):
    s = ch["page_start"] - 1
    e = min(ch["page_end"], len(pages))
    return "\n\n".join(pages[s:e])

def text_to_html(text):
    if not text.strip():
        return '<p style="color:var(--text-muted);font-style:italic">ไม่มีเนื้อหา</p>'
    
    lines = text.split('\n')
    html = []
    buffer = []

    def flush():
        if not buffer:
            return
        para = ' '.join(buffer).strip()
        buffer.clear()
        if not para or len(para) < 2:
            return
        # หัวข้อหมวด
        if re.match(r'^หมวดที่\s*\d+', para):
            html.append(f'<h2>{para}</h2>')
        # หัวข้อรอง
        elif re.match(r'^\d+\.\s+\S', para) and len(para) < 100:
            html.append(f'<h3>{para}</h3>')
        elif re.match(r'^\d+\.\d+\s+\S', para) and len(para) < 100:
            html.append(f'<h4>{para}</h4>')
        else:
            html.append(f'<p>{para}</p>')

    for line in lines:
        line = line.strip()
        if not line:
            flush()
        else:
            buffer.append(line)
    flush()
    return '\n'.join(html)


PAGE_HTML = """<!DOCTYPE html>
<html lang="th" data-lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title_th} — Akara Resources</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body>
  <header class="topbar">
    <div class="topbar-left">
      <a href="index.html" class="back-btn">
        <svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M9 2L4 7L9 12"/></svg>
        <span class="text-th">สารบัญ</span><span class="text-en">Contents</span>
      </a>
      <img src="../assets/Akara Logo.jpg" alt="Akara Resources" class="topbar-logo">
      <div class="topbar-title">
        <span class="text-th">คู่มือพนักงาน</span><span class="text-en">Employee Handbook</span>
        <span>Akara Resources Public Company Limited</span>
      </div>
    </div>
    <div class="topbar-right">
      <div class="lang-toggle">
        <button class="lang-btn active" onclick="setLang('th')">TH</button>
        <button class="lang-btn" onclick="setLang('en')">EN</button>
      </div>
    </div>
  </header>

  <main class="container">
    <button class="sidebar-toggle" onclick="toggleSidebar()">
      ☰ <span class="text-th">หัวข้อ</span><span class="text-en">Sections</span>
    </button>
    <div class="chapter-layout">
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-title text-th">หัวข้อในบทนี้</div>
        <div class="sidebar-title text-en">In this chapter</div>
        <nav class="sidebar-nav" id="sidebarNav"></nav>
      </aside>
      <article class="chapter-card">
        <div class="chapter-badge">{badge}</div>
        <h1 class="chapter-title">
          <span class="text-th">{title_th}</span>
          <span class="text-en">{title_en}</span>
        </h1>
        <div class="chapter-divider"></div>
        <div class="chapter-body" id="chapterBody">
          <div class="text-th">{content_th}</div>
          <div class="text-en">{content_en}</div>
        </div>
        <nav class="chapter-nav">{prev_link}{next_link}</nav>
      </article>
    </div>
  </main>

  <footer class="footer">
    <span class="text-th">© 2026 บริษัท อัครา รีซอร์สเซส จำกัด (มหาชน) — ฝ่ายทรัพยากรบุคคล</span>
    <span class="text-en">© 2026 Akara Resources Public Company Limited — Human Resources Department</span>
  </footer>

  <script>
    function setLang(lang) {{
      document.documentElement.setAttribute('data-lang', lang);
      document.querySelectorAll('.lang-btn').forEach(b => b.classList.remove('active'));
      document.querySelector(`.lang-btn[onclick="setLang('${{lang}}')"]`).classList.add('active');
      localStorage.setItem('hr-lang', lang);
    }}
    const saved = localStorage.getItem('hr-lang');
    if (saved) setLang(saved);

    // Build sidebar from headings
    const headings = document.querySelectorAll('.chapter-body h2, .chapter-body h3, .chapter-body h4');
    const nav = document.getElementById('sidebarNav');
    headings.forEach((h, i) => {{
      h.id = 'sec-' + i;
      const a = document.createElement('a');
      a.href = '#sec-' + i;
      a.textContent = h.textContent.substring(0, 50);
      if (h.tagName === 'H3') a.style.paddingLeft = '12px';
      if (h.tagName === 'H4') a.style.paddingLeft = '22px';
      nav.appendChild(a);
    }});

    const obs = new IntersectionObserver(entries => {{
      entries.forEach(e => {{
        if (e.isIntersecting) {{
          nav.querySelectorAll('a').forEach(a => a.classList.remove('active'));
          const active = nav.querySelector(`a[href="#${{e.target.id}}"]`);
          if (active) active.classList.add('active');
        }}
      }});
    }}, {{rootMargin: '-10% 0px -80% 0px'}});
    headings.forEach(h => obs.observe(h));

    function toggleSidebar() {{
      document.getElementById('sidebar').classList.toggle('open');
    }}
  </script>
</body>
</html>"""

TOC_ITEM = """<a href="{filename}" class="toc-item">
  <div class="toc-num">{badge}</div>
  <div class="toc-text">
    <strong class="text-th">{title_th}</strong>
    <strong class="text-en">{title_en}</strong>
  </div>
  <span class="toc-chevron">›</span>
</a>"""

def main():
    print("=" * 56)
    print("  Akara Resources — HR Policy Generator v3")
    print("=" * 56)

    if PDF_LIB is None:
        print("\n❌  pip install PyMuPDF\n"); return

    pdf_th = os.path.join(BASE_DIR, "content", "work-rules", "01_Work Rules and Regulations TH.pdf")
    pdf_en = os.path.join(BASE_DIR, "content", "work-rules", "02_Work Rules and Regulations EN.pdf")
    out_dir = os.path.join(BASE_DIR, "work-rules")
    toc_path = os.path.join(out_dir, "index.html")

    if not os.path.exists(pdf_th):
        print(f"\n❌ ไม่พบ: {pdf_th}"); return

    print(f"\n📄 อ่าน PDF TH ({pdf_th.split(chr(92))[-1]})...")
    pages_th = extract_pages(pdf_th)
    print(f"   พบ {len(pages_th)} หน้า")

    pages_en = []
    if os.path.exists(pdf_en):
        print(f"📄 อ่าน PDF EN ({pdf_en.split(chr(92))[-1]})...")
        pages_en = extract_pages(pdf_en)
        print(f"   พบ {len(pages_en)} หน้า")

    print(f"\n✍️  สร้าง HTML {len(WORK_RULES_CHAPTERS)} หน้า...")
    toc_items = []
    total = len(WORK_RULES_CHAPTERS)

    for i, ch in enumerate(WORK_RULES_CHAPTERS):
        chnum = ch["num"]
        badge = "สารจากผู้บริหาร" if chnum == 0 else f"หมวดที่ {chnum}"
        filename = "chapter-00.html" if chnum == 0 else f"chapter-{chnum:02d}.html"

        content_th = text_to_html(get_chapter_text(pages_th, ch))
        content_en = text_to_html(get_chapter_text(pages_en, ch)) if pages_en else '<p>English content</p>'

        prev_link = next_link = ""
        if i > 0:
            pch = WORK_RULES_CHAPTERS[i-1]
            pn = pch["num"]
            pf = "chapter-00.html" if pn == 0 else f"chapter-{pn:02d}.html"
            prev_link = f'<a href="{pf}">← {pch["title_th"][:35]}</a>'
        if i < total - 1:
            nch = WORK_RULES_CHAPTERS[i+1]
            nn = nch["num"]
            nf = "chapter-00.html" if nn == 0 else f"chapter-{nn:02d}.html"
            next_link = f'<a href="{nf}" class="next">{nch["title_th"][:35]} →</a>'

        html = PAGE_HTML.format(
            title_th=ch["title_th"], title_en=ch["title_en"],
            badge=badge, content_th=content_th, content_en=content_en,
            prev_link=prev_link, next_link=next_link
        )
        out_path = os.path.join(out_dir, filename)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  ✅ {filename} (หน้า {ch['page_start']}–{ch['page_end']})")

        toc_items.append(TOC_ITEM.format(
            filename=filename, badge=badge,
            title_th=ch["title_th"], title_en=ch["title_en"]
        ))

    # อัปเดต TOC
    with open(toc_path, "r", encoding="utf-8") as f:
        toc_html = f.read()

    # ลบ TOC เดิมแล้วใส่ใหม่
    toc_html = re.sub(r'<!-- WORK_RULES_TOC_PLACEHOLDER -->.*?(?=\s*</div>)', 
                      "\n".join(toc_items), toc_html, flags=re.DOTALL)
    # ถ้า placeholder ยังอยู่
    if "WORK_RULES_TOC_PLACEHOLDER" in toc_html:
        toc_html = toc_html.replace("<!-- WORK_RULES_TOC_PLACEHOLDER -->", "\n".join(toc_items))

    with open(toc_path, "w", encoding="utf-8") as f:
        f.write(toc_html)

    print(f"\n✅ อัปเดต TOC → work-rules/index.html")
    print("\n" + "="*56)
    print("  เสร็จ! รัน: git add . && git commit -m 'update chapters' && git push")
    print("="*56 + "\n")

if __name__ == "__main__":
    main()
