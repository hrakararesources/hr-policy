#!/usr/bin/env python3
"""Akara Resources — HR Policy Generator v6 (Word .docx)"""
import os, re
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

WORK_RULES_CHAPTERS = [
    {"num": 0,  "title_th": "สารจากผู้บริหาร",       "title_en": "Message from Management",
     "start": "สารจากผู้บริหาร", "end": "หมวดที่ 1"},
    {"num": 1,  "title_th": "บททั่วไป",               "title_en": "General Provisions",
     "start": "หมวดที่ 1",       "end": "หมวดที่ 2"},
    {"num": 2,  "title_th": "วิสัยทัศน์องค์กร เสาหลักกลยุทธ์และนโยบายองค์กร",
     "title_en": "Corporate Vision, Strategic Pillars and Policies",
     "start": "หมวดที่ 2",       "end": "หมวดที่ 3"},
    {"num": 3,  "title_th": "การว่าจ้าง",              "title_en": "Employment",
     "start": "หมวดที่ 3",       "end": "หมวดที่ 4"},
    {"num": 4,  "title_th": "การสับเปลี่ยนหน้าที่ การโยกย้าย และการปฏิบัติงานที่บ้าน",
     "title_en": "Job Rotation, Transfer and Work from Home",
     "start": "หมวดที่ 4",       "end": "หมวดที่ 5"},
    {"num": 5,  "title_th": "วันทำงาน เวลาทำงานปกติ และเวลาพัก",
     "title_en": "Working Days, Hours and Rest Periods",
     "start": "หมวดที่ 5",       "end": "หมวดที่ 6"},
    {"num": 6,  "title_th": "วันหยุด และหลักเกณฑ์วันหยุด",
     "title_en": "Holidays and Holiday Criteria",
     "start": "หมวดที่ 6",       "end": "หมวดที่ 7"},
    {"num": 7,  "title_th": "หลักเกณฑ์การทำงานล่วงเวลา และการทำงานในวันหยุด",
     "title_en": "Overtime and Holiday Work Criteria",
     "start": "หมวดที่ 7",       "end": "หมวดที่ 8"},
    {"num": 8,  "title_th": "การจ่ายค่าจ้าง ค่าล่วงเวลา และค่าทำงานในวันหยุด",
     "title_en": "Wages, Overtime Pay and Holiday Pay",
     "start": "หมวดที่ 8",       "end": "หมวดที่ 9"},
    {"num": 9,  "title_th": "วันลา และหลักเกณฑ์การลา",
     "title_en": "Leave and Leave Criteria",
     "start": "หมวดที่ 9",       "end": "หมวดที่ 10"},
    {"num": 10, "title_th": "วินัย และโทษทางวินัย",
     "title_en": "Discipline and Disciplinary Penalties",
     "start": "หมวดที่ 10",      "end": "หมวดที่ 11"},
    {"num": 11, "title_th": "การร้องทุกข์",
     "title_en": "Grievance Procedure",
     "start": "หมวดที่ 11",      "end": "หมวดที่ 12"},
    {"num": 12, "title_th": "การพ้นสภาพการเป็นพนักงาน",
     "title_en": "Termination of Employment",
     "start": "หมวดที่ 12",      "end": "หมวดที่ 13"},
    {"num": 13, "title_th": "สวัสดิการและสิทธิประโยชน์อื่น ๆ",
     "title_en": "Welfare and Other Benefits",
     "start": "หมวดที่ 13",      "end": "หมวดที่ 14"},
    {"num": 14, "title_th": "ข้อกำหนดทั่วไปและการบังคับใช้",
     "title_en": "General Provisions and Enforcement",
     "start": "หมวดที่ 14",      "end": None},
]

SKIP_TEXTS = [
    'สารบัญ', 'ระเบียบข้อบังคับเกี่ยวกับการทำงาน',
    'บริษัท อัครา รีซอร์สเซส จำกัด (มหาชน)',
    'ฉบับ เดือนมีนาคม', 'ประเภทกิจการ', 'ทำเหมืองแร่',
    'ห้ามทำสำเนา', 'AKR-', 'Rev.',
]

SIZE_H1 = 228600  # หมวดที่ X
SIZE_BODY = 190500

def para_to_html(para):
    """แปลง paragraph เป็น HTML ตาม bold/size"""
    t = para.text.strip()
    if not t:
        return None
    if any(s in t for s in SKIP_TEXTS):
        return None

    is_bold = any(run.bold for run in para.runs if run.text.strip())
    size = None
    for run in para.runs:
        if run.text.strip() and run.font.size:
            size = run.font.size
            break

    style = para.style.name

    # หมวดที่ X — ตัวใหญ่ bold
    if size and size >= SIZE_H1 and is_bold:
        if re.match(r'^หมวดที่\s*\d+', t):
            return f'<h2 class="chapter-section">{t}</h2>'
        else:
            return f'<h2>{t}</h2>'

    # หัวข้อ bold ขนาดปกติ
    if is_bold and style == 'Normal':
        return f'<h3>{t}</h3>'

    # List item bold = หัวข้อย่อย
    if is_bold and style == 'List Paragraph':
        # ตรวจว่าเป็นข้อเลข 1. 2. 3.
        if re.match(r'^\d+\.', t):
            return f'<h4>{t}</h4>'
        return f'<h4>{t}</h4>'

    # List item ไม่ bold = รายการ
    if style == 'List Paragraph':
        return f'<li>{t}</li>'

    # Normal ไม่ bold = paragraph
    return f'<p>{t}</p>'

def wrap_lists(html_parts):
    """ห่อ <li> ด้วย <ul>"""
    result = []
    in_list = False
    for item in html_parts:
        if item and item.startswith('<li>'):
            if not in_list:
                result.append('<ul>')
                in_list = True
            result.append(item)
        else:
            if in_list:
                result.append('</ul>')
                in_list = False
            if item:
                result.append(item)
    if in_list:
        result.append('</ul>')
    return result

def extract_chapter(paragraphs, ch):
    """ดึง paragraphs ของหมวดที่ต้องการ"""
    start_key = ch["start"]
    end_key = ch["end"]
    collecting = False
    result = []

    for para in paragraphs:
        t = para.text.strip()
        if not t:
            continue

        # เริ่มเก็บเมื่อเจอ start_key
        if not collecting:
            if t == start_key or t.startswith(start_key):
                collecting = True
                # ถ้าเป็น "สารจากผู้บริหาร" ไม่ต้อง include บรรทัดนี้
                if ch["num"] != 0:
                    result.append(para)
                continue

        if collecting:
            # หยุดเมื่อเจอ end_key
            if end_key and (t == end_key or t.startswith(end_key)):
                break
            result.append(para)

    return result

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
      width:270px; flex-shrink:0; position:sticky; top:60px;
      height:calc(100vh - 60px); overflow-y:auto;
      background:#fff; border-right:1.5px solid var(--gray-2); padding:12px 8px;
    }}
    .sidebar-fixed::-webkit-scrollbar {{ width:4px; }}
    .sidebar-fixed::-webkit-scrollbar-thumb {{ background:var(--gray-3); border-radius:2px; }}
    .sidebar-home {{
      display:flex; align-items:center; gap:6px; padding:7px 10px;
      font-size:0.8rem; color:var(--text-muted); text-decoration:none;
      border-radius:6px; margin-bottom:2px;
    }}
    .sidebar-home:hover {{ background:var(--gray-1); color:var(--blue); }}
    .sidebar-section {{
      font-size:0.68rem; font-weight:700; letter-spacing:0.1em;
      text-transform:uppercase; color:var(--text-muted);
      padding:10px 10px 4px; display:block;
    }}
    .sidebar-chapter {{
      display:flex; align-items:flex-start; gap:8px; padding:8px 10px;
      border-radius:8px; text-decoration:none; color:var(--text-muted);
      font-size:0.8rem; line-height:1.35; transition:all 0.15s; margin-bottom:1px;
    }}
    .sidebar-chapter:hover {{ background:var(--blue-light); color:var(--blue); }}
    .sidebar-chapter.active {{ background:var(--blue); color:#fff; }}
    .sidebar-chapter.active .sidebar-num {{ background:rgba(255,255,255,0.25); color:#fff; }}
    .sidebar-num {{
      min-width:26px; height:26px; border-radius:6px;
      background:var(--blue-light); color:var(--blue);
      font-size:0.7rem; font-weight:700;
      display:flex; align-items:center; justify-content:center; flex-shrink:0; margin-top:1px;
    }}
    .sidebar-label {{ flex:1; }}
    .main-content {{ flex:1; padding:32px 32px 60px; min-width:0; }}
    .chapter-card {{ max-width:740px; margin:0 auto; }}
    .chapter-body h2.chapter-section {{ color:var(--blue); font-size:1.1rem; margin:28px 0 8px; }}
    .chapter-body h2 {{ color:var(--dark); font-size:1.15rem; margin:24px 0 8px; }}
    .chapter-body h3 {{ color:var(--blue); font-size:1rem; margin:20px 0 6px; border-left:3px solid var(--blue); padding-left:10px; }}
    .chapter-body h4 {{ color:var(--dark); font-size:0.95rem; margin:14px 0 4px; font-weight:600; }}
    .chapter-body p {{ color:var(--text); font-size:0.95rem; margin-bottom:0.8em; line-height:1.75; }}
    .chapter-body ul {{ padding-left:1.4em; margin-bottom:1em; }}
    .chapter-body li {{ font-size:0.95rem; margin-bottom:0.4em; color:var(--text); line-height:1.7; }}
    @media(max-width:768px) {{
      .sidebar-fixed {{ display:none; }}
      .sidebar-fixed.open {{ display:block; position:fixed; top:60px; left:0; z-index:200; height:calc(100vh - 60px); box-shadow:4px 0 20px rgba(0,0,0,0.15); }}
      .mob-menu {{ display:flex !important; }}
    }}
  </style>
</head>
<body>
  <header class="topbar">
    <div class="topbar-left">
      <button class="mob-menu" onclick="toggleSidebar()" style="display:none;background:rgba(255,255,255,0.15);border:none;color:#fff;padding:6px 10px;border-radius:6px;cursor:pointer;font-size:1rem;margin-right:8px;">☰</button>
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
{content_th}
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
    print("  Akara Resources — HR Policy Generator v6 (docx)")
    print("="*56)

    docx_th = os.path.join(BASE_DIR, "content", "work-rules", "01_Work Rules and Regulations TH.docx")
    out_dir = os.path.join(BASE_DIR, "work-rules")
    toc_path = os.path.join(out_dir, "index.html")

    if not os.path.exists(docx_th):
        print(f"\n❌ ไม่พบ: {docx_th}")
        print("   วางไฟล์ .docx ใน content/work-rules/ แล้วลองใหม่")
        return

    print(f"\n📄 อ่าน Word TH...")
    doc_th = Document(docx_th)
    paragraphs_th = doc_th.paragraphs
    print(f"   {len(paragraphs_th)} paragraphs")

    print(f"\n✍️  สร้าง HTML {len(WORK_RULES_CHAPTERS)} หน้า...")
    toc_items = []
    total = len(WORK_RULES_CHAPTERS)

    for i, ch in enumerate(WORK_RULES_CHAPTERS):
        chnum = ch["num"]
        badge = "สารจากผู้บริหาร" if chnum == 0 else f"หมวดที่ {chnum}"
        filename = "chapter-00.html" if chnum == 0 else f"chapter-{chnum:02d}.html"

        # ดึง paragraphs ของหมวดนี้
        ch_paras = extract_chapter(paragraphs_th, ch)

        # แปลงเป็น HTML
        html_parts = [para_to_html(p) for p in ch_paras]
        html_parts = [h for h in html_parts if h]
        html_parts = wrap_lists(html_parts)
        content_th = "\n".join(html_parts)

        if not content_th.strip():
            content_th = '<p style="color:var(--text-muted)">ไม่มีเนื้อหา</p>'

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
            badge=badge, content_th=content_th,
            prev_link=prev_link, next_link=next_link,
            sidebar_links=sidebar_links
        )

        out_path = os.path.join(out_dir, filename)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  ✅ {filename} ({len(ch_paras)} paragraphs)")

        toc_items.append(TOC_ITEM.format(
            filename=filename, badge=badge,
            title_th=ch["title_th"], title_en=ch["title_en"]
        ))

    # อัปเดต TOC — replace ทั้งหมดในครั้งเดียว
    with open(toc_path, "r", encoding="utf-8") as f:
        toc_html = f.read()

    toc_block = "\n".join(toc_items)
    toc_html = re.sub(
        r'<div class="toc-list"[^>]*>.*?</div>',
        f'<div class="toc-list" id="tocList">\n{toc_block}\n</div>',
        toc_html, flags=re.DOTALL
    )

    with open(toc_path, "w", encoding="utf-8") as f:
        f.write(toc_html)

    print(f"\n✅ อัปเดต TOC")
    print("\n" + "="*56)
    print("  เสร็จ! git add . && git commit -m 'v6 docx' && git push")
    print("="*56 + "\n")

if __name__ == "__main__":
    main()
