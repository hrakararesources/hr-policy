#!/usr/bin/env python3
"""Akara Resources — HR Policy HTML Generator v4 (sidebar nav)"""

import os, re

try:
    import fitz
    PDF_LIB = "pymupdf"
except ImportError:
    PDF_LIB = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

WORK_RULES_CHAPTERS = [
    {"num": 0,  "title_th": "สารจากผู้บริหาร",       "title_en": "Message from Management",        "page_start": 2, "page_end": 2},
    {"num": 1,  "title_th": "บททั่วไป",               "title_en": "General Provisions",              "page_start": 3, "page_end": 6},
    {"num": 2,  "title_th": "วิสัยทัศน์องค์กร เสาหลักกลยุทธ์และนโยบายองค์กร", "title_en": "Corporate Vision, Strategic Pillars and Policies", "page_start": 7, "page_end": 9},
    {"num": 3,  "title_th": "การว่าจ้าง",              "title_en": "Employment",                      "page_start": 10, "page_end": 11},
    {"num": 4,  "title_th": "การสับเปลี่ยนหน้าที่ การโยกย้าย และการปฏิบัติงานที่บ้าน", "title_en": "Job Rotation, Transfer and Work from Home", "page_start": 12, "page_end": 13},
    {"num": 5,  "title_th": "วันทำงาน เวลาทำงานปกติ และเวลาพัก", "title_en": "Working Days, Hours and Rest Periods", "page_start": 14, "page_end": 16},
    {"num": 6,  "title_th": "วันหยุด และหลักเกณฑ์วันหยุด", "title_en": "Holidays and Holiday Criteria", "page_start": 17, "page_end": 19},
    {"num": 7,  "title_th": "หลักเกณฑ์การทำงานล่วงเวลา และการทำงานในวันหยุด", "title_en": "Overtime and Holiday Work", "page_start": 20, "page_end": 21},
    {"num": 8,  "title_th": "การจ่ายค่าจ้าง ค่าล่วงเวลา และค่าทำงานในวันหยุด", "title_en": "Wages, Overtime and Holiday Pay", "page_start": 22, "page_end": 26},
    {"num": 9,  "title_th": "วันลา และหลักเกณฑ์การลา", "title_en": "Leave and Leave Criteria",        "page_start": 27, "page_end": 33},
    {"num": 10, "title_th": "วินัย และโทษทางวินัย",    "title_en": "Discipline and Penalties",        "page_start": 34, "page_end": 44},
    {"num": 11, "title_th": "การร้องทุกข์",             "title_en": "Grievance Procedure",             "page_start": 45, "page_end": 46},
    {"num": 12, "title_th": "การพ้นสภาพการเป็นพนักงาน","title_en": "Termination of Employment",       "page_start": 47, "page_end": 51},
    {"num": 13, "title_th": "สวัสดิการและสิทธิประโยชน์อื่น ๆ", "title_en": "Welfare and Other Benefits", "page_start": 52, "page_end": 53},
    {"num": 14, "title_th": "ข้อกำหนดทั่วไปและการบังคับใช้", "title_en": "General Provisions and Enforcement", "page_start": 54, "page_end": 55},
]

JUNK = [
    r'ห้ามทำสำเนา[^\n]*', r'ห้ามท า[^\n]*',
    r'AKR-DCC[^\n]*', r'AKR-OHR[^\n]*',
    r'Support Document[^\n]*',
    r'Document No\.[^\n]*', r'Document Title:[^\n]*',
    r'Revision No\.[^\n]*', r'Effective Date[^\n]*',
    r'Page No\.[^\n]*',
    r'Work Rules and Regulations\s*', r'Welfare and Benefits\s*',
    r'^\d+\s*of\s*\d+\s*$', r'^Rev\.\d+.*$',
    r'^00\s*$', r'^16-Apr-2026\s*$',
]

def clean(text):
    for p in JUNK:
        text = re.sub(p, '', text, flags=re.MULTILINE)
    return re.sub(r'\n{3,}', '\n\n', text).strip()

def extract_pages(pdf_path):
    doc = fitz.open(pdf_path)
    pages = [clean(p.get_text("text")) for p in doc]
    doc.close()
    return pages

def get_chapter_text(pages, ch):
    s = ch["page_start"] - 1
    e = min(ch["page_end"], len(pages))
    return "\n\n".join(pages[s:e])

def text_to_html(text):
    if not text.strip():
        return '<p style="color:var(--text-muted)">ไม่มีเนื้อหา</p>'
    html, buffer = [], []
    def flush():
        if not buffer: return
        para = ' '.join(buffer).strip()
        buffer.clear()
        if not para or len(para) < 2: return
        if re.match(r'^หมวดที่\s*\d+', para):
            html.append(f'<h2>{para}</h2>')
        elif re.match(r'^\d+\.\s+\S', para) and len(para) < 120:
            html.append(f'<h3>{para}</h3>')
        elif re.match(r'^\d+\.\d+\s+\S', para) and len(para) < 120:
            html.append(f'<h4>{para}</h4>')
        else:
            html.append(f'<p>{para}</p>')
    for line in text.split('\n'):
        line = line.strip()
        if not line: flush()
        else: buffer.append(line)
    flush()
    return '\n'.join(html)

def build_sidebar(chapters, current_file):
    links = []
    for c in chapters:
        cn = c["num"]
        cf = "chapter-00.html" if cn == 0 else f"chapter-{cn:02d}.html"
        num_label = "00" if cn == 0 else f"{cn:02d}"
        active = ' active' if cf == current_file else ''
        links.append(
            f'<a href="{cf}" class="sidebar-chapter{active}">'
            f'<span class="sidebar-num">{num_label}</span>'
            f'<span class="sidebar-label">'
            f'<span class="text-th">{c["title_th"]}</span>'
            f'<span class="text-en">{c["title_en"]}</span>'
            f'</span></a>'
        )
    return "\n          ".join(links)

PAGE_HTML = '''<!DOCTYPE html>
<html lang="th" data-lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title_th} — Akara Resources</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="../assets/style.css">
  <style>
    .chapter-layout {{ grid-template-columns: 280px 1fr; }}
    .sidebar-chapter {{
      display: flex; align-items: flex-start; gap: 10px;
      padding: 10px 12px; border-radius: 8px;
      text-decoration: none; color: var(--text-muted);
      font-size: 0.82rem; line-height: 1.35;
      transition: all 0.2s; border: 1.5px solid transparent;
    }}
    .sidebar-chapter:hover {{ background: var(--blue-light); color: var(--blue); }}
    .sidebar-chapter.active {{
      background: var(--blue); color: #fff;
      border-color: var(--blue);
    }}
    .sidebar-chapter.active .sidebar-num {{ background: rgba(255,255,255,0.2); color: #fff; }}
    .sidebar-num {{
      min-width: 28px; height: 28px; border-radius: 7px;
      background: var(--blue-light); color: var(--blue);
      font-size: 0.72rem; font-weight: 700;
      display: flex; align-items: center; justify-content: center;
      flex-shrink: 0; margin-top: 1px;
    }}
    .sidebar-label {{ flex: 1; }}
    .sidebar-section-label {{
      font-size: 0.68rem; font-weight: 700; letter-spacing: 0.1em;
      text-transform: uppercase; color: var(--text-muted);
      padding: 12px 12px 4px; display: block;
    }}
  </style>
</head>
<body>
  <header class="topbar">
    <div class="topbar-left">
      <img src="../assets/Akara Logo.jpg" alt="Akara Resources" class="topbar-logo">
      <div class="topbar-title">
        <span class="text-th">คู่มือพนักงาน</span><span class="text-en">Employee Handbook</span>
        <span>Akara Resources Public Company Limited</span>
      </div>
    </div>
    <div class="topbar-right">
      <div class="lang-toggle">
        <button class="lang-btn active" onclick="setLang(\'th\')">TH</button>
        <button class="lang-btn" onclick="setLang(\'en\')">EN</button>
      </div>
    </div>
  </header>

  <div style="display:flex; min-height:calc(100vh - 60px);">
    <aside class="sidebar" id="sidebar" style="width:280px;flex-shrink:0;position:sticky;top:60px;height:calc(100vh - 60px);overflow-y:auto;background:#fff;border-right:1px solid var(--gray-2);padding:16px 8px;">
      <a href="../index.html" style="display:flex;align-items:center;gap:6px;padding:8px 12px;font-size:0.8rem;color:var(--text-muted);text-decoration:none;margin-bottom:4px;">
        ← <span class="text-th">หน้าหลัก</span><span class="text-en">Home</span>
      </a>
      <a href="index.html" style="display:flex;align-items:center;gap:6px;padding:8px 12px;font-size:0.8rem;color:var(--blue);text-decoration:none;font-weight:600;margin-bottom:8px;">
        ☰ <span class="text-th">สารบัญ</span><span class="text-en">Contents</span>
      </a>
      <span class="sidebar-section-label text-th">ระเบียบข้อบังคับการทำงาน</span>
      <span class="sidebar-section-label text-en">Work Rules & Regulations</span>
      <nav style="display:flex;flex-direction:column;gap:2px;">
          {sidebar_links}
      </nav>
    </aside>

    <main style="flex:1;padding:32px 28px 60px;min-width:0;">
      <button class="sidebar-toggle" onclick="toggleSidebar()" style="display:none;">☰</button>
      <article class="chapter-card" style="max-width:760px;margin:0 auto;">
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
    </main>
  </div>

  <footer class="footer">
    <span class="text-th">© 2026 บริษัท อัครา รีซอร์สเซส จำกัด (มหาชน) — ฝ่ายทรัพยากรบุคคล</span>
    <span class="text-en">© 2026 Akara Resources Public Company Limited — Human Resources Department</span>
  </footer>

  <script>
    function setLang(lang) {{
      document.documentElement.setAttribute(\'data-lang\', lang);
      document.querySelectorAll(\'.lang-btn\').forEach(b => b.classList.remove(\'active\'));
      document.querySelector(`.lang-btn[onclick="setLang(\'${{lang}}\')"]`).classList.add(\'active\');
      localStorage.setItem(\'hr-lang\', lang);
    }}
    const saved = localStorage.getItem(\'hr-lang\');
    if (saved) setLang(saved);

    function toggleSidebar() {{
      document.getElementById(\'sidebar\').classList.toggle(\'open\');
    }}
  </script>
</body>
</html>'''

TOC_ITEM = '''<a href="{filename}" class="toc-item">
  <div class="toc-num">{badge}</div>
  <div class="toc-text">
    <strong class="text-th">{title_th}</strong>
    <strong class="text-en">{title_en}</strong>
  </div>
  <span class="toc-chevron">›</span>
</a>'''

def main():
    print("="*56)
    print("  Akara Resources — HR Policy Generator v4")
    print("="*56)
    if PDF_LIB is None:
        print("\n❌  pip install PyMuPDF\n"); return

    pdf_th = os.path.join(BASE_DIR, "content", "work-rules", "01_Work Rules and Regulations TH.pdf")
    pdf_en = os.path.join(BASE_DIR, "content", "work-rules", "02_Work Rules and Regulations EN.pdf")
    out_dir = os.path.join(BASE_DIR, "work-rules")
    toc_path = os.path.join(out_dir, "index.html")

    if not os.path.exists(pdf_th):
        print(f"\n❌ ไม่พบ: {pdf_th}"); return

    print(f"\n📄 อ่าน PDF TH...")
    pages_th = extract_pages(pdf_th)
    print(f"   {len(pages_th)} หน้า")

    pages_en = []
    if os.path.exists(pdf_en):
        print(f"📄 อ่าน PDF EN...")
        pages_en = extract_pages(pdf_en)
        print(f"   {len(pages_en)} หน้า")

    print(f"\n✍️  สร้าง HTML {len(WORK_RULES_CHAPTERS)} หน้า...")
    toc_items = []
    total = len(WORK_RULES_CHAPTERS)

    for i, ch in enumerate(WORK_RULES_CHAPTERS):
        chnum = ch["num"]
        badge = "สารจากผู้บริหาร" if chnum == 0 else f"หมวดที่ {chnum}"
        filename = "chapter-00.html" if chnum == 0 else f"chapter-{chnum:02d}.html"

        content_th = text_to_html(get_chapter_text(pages_th, ch))
        content_en = text_to_html(get_chapter_text(pages_en, ch)) if pages_en else '<p>English content</p>'
        sidebar_links = build_sidebar(WORK_RULES_CHAPTERS, filename)

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
            prev_link=prev_link, next_link=next_link,
            sidebar_links=sidebar_links
        )
        out_path = os.path.join(out_dir, filename)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  ✅ {filename}")

        toc_items.append(TOC_ITEM.format(
            filename=filename, badge=badge,
            title_th=ch["title_th"], title_en=ch["title_en"]
        ))

    with open(toc_path, "r", encoding="utf-8") as f:
        toc_html = f.read()
    toc_block = "\n".join(toc_items)
    if "WORK_RULES_TOC_PLACEHOLDER" in toc_html:
        toc_html = toc_html.replace("<!-- WORK_RULES_TOC_PLACEHOLDER -->", toc_block)
    else:
        toc_html = re.sub(r'<div class="toc-list"[^>]*>.*?</div>',
                          f'<div class="toc-list" id="tocList">{toc_block}</div>',
                          toc_html, flags=re.DOTALL)
    with open(toc_path, "w", encoding="utf-8") as f:
        f.write(toc_html)

    print(f"\n✅ อัปเดต TOC")
    print("\n" + "="*56)
    print("  เสร็จ! git add . && git commit -m 'v4 sidebar' && git push")
    print("="*56 + "\n")

if __name__ == "__main__":
    main()
