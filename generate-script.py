#!/usr/bin/env python3
"""Akara Resources — HR Policy Generator v5 (correct page ranges + sara am fix)"""
import os, re

try:
    import fitz
    PDF_LIB = "pymupdf"
except ImportError:
    PDF_LIB = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

WORK_RULES_CHAPTERS = [
    {"num": 0,  "title_th": "สารจากผู้บริหาร",       "title_en": "Message from Management",               "ps": 3,  "pe": 3},
    {"num": 1,  "title_th": "บททั่วไป",               "title_en": "General Provisions",                     "ps": 4,  "pe": 7},
    {"num": 2,  "title_th": "วิสัยทัศน์องค์กร เสาหลักกลยุทธ์และนโยบายองค์กร", "title_en": "Corporate Vision, Strategic Pillars and Policies", "ps": 8, "pe": 10},
    {"num": 3,  "title_th": "การว่าจ้าง",              "title_en": "Employment",                             "ps": 11, "pe": 12},
    {"num": 4,  "title_th": "การสับเปลี่ยนหน้าที่ การโยกย้าย และการปฏิบัติงานที่บ้าน", "title_en": "Job Rotation, Transfer and Work from Home", "ps": 13, "pe": 14},
    {"num": 5,  "title_th": "วันทำงาน เวลาทำงานปกติ และเวลาพัก", "title_en": "Working Days, Hours and Rest Periods", "ps": 15, "pe": 17},
    {"num": 6,  "title_th": "วันหยุด และหลักเกณฑ์วันหยุด", "title_en": "Holidays and Holiday Criteria",    "ps": 18, "pe": 20},
    {"num": 7,  "title_th": "หลักเกณฑ์การทำงานล่วงเวลา และการทำงานในวันหยุด", "title_en": "Overtime and Holiday Work Criteria", "ps": 21, "pe": 22},
    {"num": 8,  "title_th": "การจ่ายค่าจ้าง ค่าล่วงเวลา และค่าทำงานในวันหยุด", "title_en": "Wages, Overtime Pay and Holiday Pay", "ps": 23, "pe": 27},
    {"num": 9,  "title_th": "วันลา และหลักเกณฑ์การลา", "title_en": "Leave and Leave Criteria",             "ps": 28, "pe": 34},
    {"num": 10, "title_th": "วินัย และโทษทางวินัย",    "title_en": "Discipline and Disciplinary Penalties", "ps": 35, "pe": 45},
    {"num": 11, "title_th": "การร้องทุกข์",             "title_en": "Grievance Procedure",                   "ps": 46, "pe": 47},
    {"num": 12, "title_th": "การพ้นสภาพการเป็นพนักงาน","title_en": "Termination of Employment",             "ps": 48, "pe": 52},
    {"num": 13, "title_th": "สวัสดิการและสิทธิประโยชน์อื่น ๆ", "title_en": "Welfare and Other Benefits",  "ps": 53, "pe": 54},
    {"num": 14, "title_th": "ข้อกำหนดทั่วไปและการบังคับใช้", "title_en": "General Provisions and Enforcement", "ps": 55, "pe": 56},
]

# บรรทัดที่ต้องกรองออก
SKIP = [
    'Support Document', 'AKR-OHR', 'AKR-DCC', 'Document No',
    'Document Title', 'Revision No', 'Effective Date', 'Effective date',
    'Page No', 'Work Rules and Regulations', 'Welfare and Benefits',
    'ห้ามท า', 'ห้ามทำ', 'QMR', '16-Apr-2026', '20/Jun/2025',
]

def fix_spaces(text):
    """แก้ปัญหาช่องว่างในคำภาษาไทย เช่น ท า→ทำ, ส า→สำ, จ า→จำ"""
    # แก้ pattern: พยัญชนะ + space + า (สระอา/อำที่ถูกแยก)
    fixes = [
        (r'ท า', 'ทำ'), (r'ส า', 'สำ'), (r'จ า', 'จำ'), (r'ก า', 'กำ'),
        (r'ค า', 'คำ'), (r'ข า', 'ขำ'), (r'ง า', 'งำ'), (r'ด า', 'ดำ'),
        (r'ต า', 'ตำ'), (r'น า', 'นำ'), (r'บ า', 'บำ'), (r'ป า', 'ปำ'),
        (r'ผ า', 'ผำ'), (r'ฝ า', 'ฝำ'), (r'พ า', 'พำ'), (r'ฟ า', 'ฟำ'),
        (r'ภ า', 'ภำ'), (r'ม า', 'มำ'), (r'ย า', 'ยำ'), (r'ร า', 'รำ'),
        (r'ล า', 'ลำ'), (r'ว า', 'วำ'), (r'ห า', 'หำ'), (r'อ า', 'อำ'),
        (r'ช า', 'ชำ'), (r'ซ า', 'ซำ'), (r'ฉ า', 'ฉำ'), (r'ฃ า', 'ฃำ'),
        (r'ท ำ', 'ทำ'), (r'ส ำ', 'สำ'), (r'จ ำ', 'จำ'), (r'ก ำ', 'กำ'),
        # แก้ช่องว่างใน ำ
        (r' ำ', 'ำ'), (r'า ', 'า'),
    ]
    for old, new in fixes:
        text = text.replace(old, new)
    return text

def clean_line(line):
    """ตรวจว่าควรเก็บบรรทัดนี้ไหม"""
    line = line.strip()
    if not line:
        return None
    if any(s in line for s in SKIP):
        return None
    if re.match(r'^\d+\s+of\s+\d+$', line):
        return None
    if re.match(r'^00$', line):
        return None
    if re.match(r'^\d{1,2}-\w{3}-\d{4}$', line):
        return None
    return fix_spaces(line)

def extract_pages(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []
    for page in doc:
        raw_lines = page.get_text("text").split('\n')
        clean = [clean_line(l) for l in raw_lines]
        clean = [l for l in clean if l]
        pages.append('\n'.join(clean))
    doc.close()
    return pages

def get_chapter_text(pages, ch):
    s = ch["ps"] - 1
    e = min(ch["pe"], len(pages))
    return '\n\n'.join(pages[s:e])

def text_to_html(text):
    if not text.strip():
        return '<p style="color:var(--text-muted)">ไม่มีเนื้อหา</p>'
    html, buf = [], []
    def flush():
        if not buf: return
        para = ' '.join(buf).strip()
        buf.clear()
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
        else: buf.append(line)
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
    body {{ display:flex; flex-direction:column; min-height:100vh; }}
    .page-wrap {{ display:flex; flex:1; }}
    .sidebar-fixed {{
      width:270px; flex-shrink:0;
      position:sticky; top:60px;
      height:calc(100vh - 60px);
      overflow-y:auto;
      background:#fff;
      border-right:1.5px solid var(--gray-2);
      padding:12px 8px;
    }}
    .sidebar-fixed::-webkit-scrollbar {{ width:4px; }}
    .sidebar-fixed::-webkit-scrollbar-thumb {{ background:var(--gray-3); border-radius:2px; }}
    .sidebar-home {{
      display:flex; align-items:center; gap:6px;
      padding:7px 10px; font-size:0.8rem;
      color:var(--text-muted); text-decoration:none;
      border-radius:6px; margin-bottom:2px;
    }}
    .sidebar-home:hover {{ background:var(--gray-1); color:var(--blue); }}
    .sidebar-section {{
      font-size:0.68rem; font-weight:700; letter-spacing:0.1em;
      text-transform:uppercase; color:var(--text-muted);
      padding:10px 10px 4px; display:block;
    }}
    .sidebar-chapter {{
      display:flex; align-items:flex-start; gap:8px;
      padding:8px 10px; border-radius:8px;
      text-decoration:none; color:var(--text-muted);
      font-size:0.8rem; line-height:1.35;
      transition:all 0.15s; margin-bottom:1px;
    }}
    .sidebar-chapter:hover {{ background:var(--blue-light); color:var(--blue); }}
    .sidebar-chapter.active {{ background:var(--blue); color:#fff; }}
    .sidebar-chapter.active .sidebar-num {{ background:rgba(255,255,255,0.25); color:#fff; }}
    .sidebar-num {{
      min-width:26px; height:26px; border-radius:6px;
      background:var(--blue-light); color:var(--blue);
      font-size:0.7rem; font-weight:700;
      display:flex; align-items:center; justify-content:center;
      flex-shrink:0; margin-top:1px;
    }}
    .sidebar-label {{ flex:1; }}
    .main-content {{ flex:1; padding:32px 32px 60px; min-width:0; }}
    .chapter-card {{ max-width:740px; margin:0 auto; }}
    @media(max-width:768px) {{
      .sidebar-fixed {{ display:none; }}
      .sidebar-fixed.open {{ display:block; position:fixed; top:60px; left:0; z-index:200; height:calc(100vh - 60px); box-shadow:4px 0 20px rgba(0,0,0,0.15); }}
      .sidebar-toggle-btn {{ display:flex !important; }}
    }}
  </style>
</head>
<body>
  <header class="topbar">
    <div class="topbar-left">
      <button class="sidebar-toggle-btn" onclick="toggleSidebar()" style="display:none;background:rgba(255,255,255,0.15);border:none;color:#fff;padding:6px 10px;border-radius:6px;cursor:pointer;font-size:1rem;margin-right:8px;">☰</button>
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

  <div class="page-wrap">
    <aside class="sidebar-fixed" id="sidebar">
      <a href="../index.html" class="sidebar-home">← <span class="text-th">หน้าหลัก</span><span class="text-en">Home</span></a>
      <a href="index.html" class="sidebar-home" style="color:var(--blue);font-weight:600;">☰ <span class="text-th">สารบัญ</span><span class="text-en">Contents</span></a>
      <span class="sidebar-section text-th">ระเบียบข้อบังคับการทำงาน</span>
      <span class="sidebar-section text-en">Work Rules & Regulations</span>
      <nav style="display:flex;flex-direction:column;gap:1px;">
          {sidebar_links}
      </nav>
    </aside>

    <main class="main-content">
      <article class="chapter-card">
        <div class="chapter-badge">{badge}</div>
        <h1 class="chapter-title">
          <span class="text-th">{title_th}</span>
          <span class="text-en">{title_en}</span>
        </h1>
        <div class="chapter-divider"></div>
        <div class="chapter-body">
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
    print("  Akara Resources — HR Policy Generator v5")
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
        with open(os.path.join(out_dir, filename), "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  ✅ {filename} (หน้า {ch['ps']}–{ch['pe']})")

        toc_items.append(TOC_ITEM.format(
            filename=filename, badge=badge,
            title_th=ch["title_th"], title_en=ch["title_en"]
        ))

    # อัปเดต TOC
    with open(toc_path, "r", encoding="utf-8") as f:
        toc_html = f.read()
    toc_block = "\n".join(toc_items)
    toc_html = re.sub(
        r'<div class="toc-list"[^>]*>.*?</div>',
        f'<div class="toc-list" id="tocList">{toc_block}</div>',
        toc_html, flags=re.DOTALL
    )
    with open(toc_path, "w", encoding="utf-8") as f:
        f.write(toc_html)

    print(f"\n✅ อัปเดต TOC")
    print("\n" + "="*56)
    print("  เสร็จ! git add . && git commit -m 'v5 fix pages and sara am' && git push")
    print("="*56 + "\n")

if __name__ == "__main__":
    main()
