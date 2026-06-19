#!/usr/bin/env python3
"""
Akara Resources — HR Policy HTML Generator v2
อ่าน PDF 2 ไฟล์ (TH + EN) แยกตามหมวด แล้วสร้าง HTML
"""

import os
import re

try:
    import fitz
    PDF_LIB = "pymupdf"
except ImportError:
    PDF_LIB = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# CONFIG — หมวดของ Work Rules (ตรวจสอบกับ PDF จริง)
# ============================================================
WORK_RULES_CHAPTERS = [
    {"num": 0,  "title_th": "สารจากผู้บริหาร",       "title_en": "Message from Management",        "start_pattern": r"สารจากผู้บริหาร"},
    {"num": 1,  "title_th": "บททั่วไป",               "title_en": "General Provisions",              "start_pattern": r"หมวดที่\s*1"},
    {"num": 2,  "title_th": "วิสัยทัศน์องค์กร เสาหลักกลยุทธ์และนโยบายองค์กร", "title_en": "Corporate Vision, Strategic Pillars and Policies", "start_pattern": r"หมวดที่\s*2"},
    {"num": 3,  "title_th": "การว่าจ้าง",              "title_en": "Employment",                      "start_pattern": r"หมวดที่\s*3"},
    {"num": 4,  "title_th": "การสับเปลี่ยนหน้าที่การปฏิบัติงาน การโยกย้าย และการปฏิบัติงานที่บ้าน", "title_en": "Job Rotation, Transfer and Work from Home", "start_pattern": r"หมวดที่\s*4"},
    {"num": 5,  "title_th": "วันทำงาน เวลาทำงานปกติและเวลาพัก", "title_en": "Working Days, Hours and Rest Periods", "start_pattern": r"หมวดที่\s*5"},
    {"num": 6,  "title_th": "วันหยุด และหลักเกณฑ์วันหยุด", "title_en": "Holidays and Holiday Criteria", "start_pattern": r"หมวดที่\s*6"},
    {"num": 7,  "title_th": "หลักเกณฑ์การทำงานล่วงเวลาและการทำงานในวันหยุด", "title_en": "Overtime and Holiday Work Criteria", "start_pattern": r"หมวดที่\s*7"},
    {"num": 8,  "title_th": "การจ่ายค่าจ้าง ค่าล่วงเวลา และค่าทำงานในวันหยุด", "title_en": "Wages, Overtime Pay and Holiday Pay", "start_pattern": r"หมวดที่\s*8"},
    {"num": 9,  "title_th": "วันลาและหลักเกณฑ์การลา", "title_en": "Leave and Leave Criteria",       "start_pattern": r"หมวดที่\s*9"},
    {"num": 10, "title_th": "วินัย และโทษทางวินัย",   "title_en": "Discipline and Disciplinary Penalties", "start_pattern": r"หมวดที่\s*10"},
    {"num": 11, "title_th": "การร้องทุกข์",            "title_en": "Grievance Procedure",             "start_pattern": r"หมวดที่\s*11"},
    {"num": 12, "title_th": "การพ้นสภาพการเป็นพนักงาน", "title_en": "Termination of Employment",    "start_pattern": r"หมวดที่\s*12"},
    {"num": 13, "title_th": "สวัสดิการและสิทธิประโยชน์อื่น ๆ", "title_en": "Welfare and Other Benefits", "start_pattern": r"หมวดที่\s*13"},
    {"num": 14, "title_th": "ข้อกำหนดทั่วไปและการบังคับใช้", "title_en": "General Provisions and Enforcement", "start_pattern": r"หมวดที่\s*14"},
]

# ============================================================
# EXTRACT PDF TEXT
# ============================================================
def extract_pages(pdf_path):
    if PDF_LIB != "pymupdf":
        return []
    doc = fitz.open(pdf_path)
    pages = []
    for page in doc:
        text = page.get_text("text")
        # ล้างเส้นส่วนท้ายกระดาษ
        text = re.sub(r'ห้ามทำสำเนา.*?\n', '', text)
        text = re.sub(r'AKR-DCC.*?\n', '', text)
        text = re.sub(r'Support Document.*?\n', '', text)
        text = re.sub(r'Document No\..*?\n', '', text)
        text = re.sub(r'Document Title:.*?\n', '', text)
        text = re.sub(r'Revision No\..*?\n', '', text)
        text = re.sub(r'Effective Date.*?\n', '', text)
        text = re.sub(r'Page No\..*?\n', '', text)
        text = re.sub(r'Work Rules and Regulations\s*\n', '', text)
        pages.append(text)
    doc.close()
    return pages


def split_by_chapters(pages, chapters):
    full_text = "\n".join(pages)
    lines = full_text.split("\n")
    
    chapter_texts = {i: [] for i in range(len(chapters))}
    current_chapter = 0
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        # ตรวจว่าเป็นจุดเริ่มต้น chapter ใหม่
        matched = False
        for i, ch in enumerate(chapters):
            if i > current_chapter and re.search(ch["start_pattern"], stripped):
                current_chapter = i
                matched = True
                break
        chapter_texts[current_chapter].append(stripped)
    
    return chapter_texts


def text_to_html(lines):
    """แปลงข้อความเป็น HTML ที่อ่านง่าย"""
    html_parts = []
    buffer = []
    
    def flush_buffer():
        if not buffer:
            return
        para = " ".join(buffer).strip()
        if not para:
            buffer.clear()
            return
        # ตรวจว่าเป็นหัวข้อหรือไม่
        if re.match(r'^(หมวดที่|บทที่|\d+\.\s+[ก-ฮA-Z]|[ก-ฮ]\.\s+)', para) and len(para) < 120:
            if re.match(r'^หมวดที่', para):
                html_parts.append(f'<h2 class="chapter-section">{para}</h2>')
            else:
                html_parts.append(f'<h3>{para}</h3>')
        elif re.match(r'^\d+\.\d+', para) and len(para) < 100:
            html_parts.append(f'<h4>{para}</h4>')
        else:
            html_parts.append(f'<p>{para}</p>')
        buffer.clear()
    
    for line in lines:
        line = line.strip()
        if not line:
            flush_buffer()
        else:
            buffer.append(line)
    flush_buffer()
    
    return "\n".join(html_parts)


# ============================================================
# HTML TEMPLATE
# ============================================================
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
      <img src="../assets/Akara_Logo.jpg" alt="Akara Resources" class="topbar-logo">
      <div class="topbar-title">
        <span class="text-th">ระเบียบข้อบังคับการทำงาน</span>
        <span class="text-en">Work Rules & Regulations</span>
        <span>{badge}</span>
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
    <button class="sidebar-toggle" onclick="toggleSidebar()">☰ <span class="text-th">หัวข้อ</span><span class="text-en">Sections</span></button>
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

  <footer class="footer">© 2026 Akara Resources Co., Ltd. — Human Resources Department</footer>

  <script>
    function setLang(lang) {{
      document.documentElement.setAttribute('data-lang', lang);
      document.querySelectorAll('.lang-btn').forEach(b => b.classList.remove('active'));
      document.querySelector(`.lang-btn[onclick="setLang('${{lang}}')"]`).classList.add('active');
      localStorage.setItem('hr-lang', lang);
    }}
    const saved = localStorage.getItem('hr-lang');
    if (saved) setLang(saved);

    const headings = document.querySelectorAll('.chapter-body h3, .chapter-body h4');
    const nav = document.getElementById('sidebarNav');
    headings.forEach((h, i) => {{
      h.id = 'sec-' + i;
      const a = document.createElement('a');
      a.href = '#sec-' + i;
      a.textContent = h.textContent.substring(0,45);
      if (h.tagName === 'H4') a.style.paddingLeft = '16px';
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
    }}, {{rootMargin: '-20% 0px -75% 0px'}});
    headings.forEach(h => obs.observe(h));

    function toggleSidebar() {{
      document.getElementById('sidebar').classList.toggle('open');
    }}
  </script>
</body>
</html>"""

TOC_ITEM = """<a href="{filename}" class="toc-item">
  <div class="toc-num">{num}</div>
  <div class="toc-text">
    <strong class="text-th">{title_th}</strong>
    <strong class="text-en">{title_en}</strong>
    <span>{filename}</span>
  </div>
  <span class="toc-chevron">›</span>
</a>"""


# ============================================================
# MAIN
# ============================================================
def main():
    print("=" * 56)
    print("  Akara Resources — HR Policy Generator v2")
    print("=" * 56)

    if PDF_LIB is None:
        print("\n❌  pip install PyMuPDF\n"); return

    # ---- Work Rules ----
    pdf_th = os.path.join(BASE_DIR, "content", "work-rules", "01_Work Rules and Regulations TH.pdf")
    pdf_en = os.path.join(BASE_DIR, "content", "work-rules", "02_Work Rules and Regulations EN.pdf")
    out_dir = os.path.join(BASE_DIR, "work-rules")
    toc_path = os.path.join(out_dir, "index.html")

    if not os.path.exists(pdf_th):
        print(f"\n⚠️  ไม่พบ: {pdf_th}"); return

    print(f"\n📄 อ่าน PDF TH...")
    pages_th = extract_pages(pdf_th)
    chapter_texts_th = split_by_chapters(pages_th, WORK_RULES_CHAPTERS)

    chapter_texts_en = {i: [] for i in range(len(WORK_RULES_CHAPTERS))}
    if os.path.exists(pdf_en):
        print(f"📄 อ่าน PDF EN...")
        pages_en = extract_pages(pdf_en)
        chapter_texts_en = split_by_chapters(pages_en, WORK_RULES_CHAPTERS)

    print(f"\n✍️  สร้าง HTML {len(WORK_RULES_CHAPTERS)} หน้า...")
    toc_items = []
    total = len(WORK_RULES_CHAPTERS)

    for i, ch in enumerate(WORK_RULES_CHAPTERS):
        badge = "สารจากผู้บริหาร" if ch["num"] == 0 else f"หมวดที่ {ch['num']}"
        badge_en = "Management Message" if ch["num"] == 0 else f"Chapter {ch['num']}"
        filename = f"chapter-{'00' if ch['num']==0 else f'{ch[\"num\"]:02d}'}.html"
        
        content_th = text_to_html(chapter_texts_th.get(i, []))
        content_en = text_to_html(chapter_texts_en.get(i, []))
        
        if not content_th:
            content_th = f'<p style="color:var(--text-muted)">เนื้อหาหมวด {ch["num"]}</p>'
        if not content_en:
            content_en = f'<p style="color:var(--text-muted)">Content for chapter {ch["num"]}</p>'

        # prev/next
        prev_link = next_link = ""
        if i > 0:
            prev_ch = WORK_RULES_CHAPTERS[i-1]
            prev_file = f"chapter-{'00' if prev_ch['num']==0 else f'{prev_ch[\"num\"]:02d}'}.html"
            prev_link = f'<a href="{prev_file}">← {prev_ch["title_th"][:30]}</a>'
        if i < total - 1:
            next_ch = WORK_RULES_CHAPTERS[i+1]
            next_file = f"chapter-{'00' if next_ch['num']==0 else f'{next_ch[\"num\"]:02d}'}.html"
            next_link = f'<a href="{next_file}" class="next">{next_ch["title_th"][:30]} →</a>'

        html = PAGE_HTML.format(
            title_th=ch["title_th"], title_en=ch["title_en"],
            badge=badge, content_th=content_th, content_en=content_en,
            prev_link=prev_link, next_link=next_link
        )
        out_path = os.path.join(out_dir, filename)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  ✅ {filename}")

        toc_items.append(TOC_ITEM.format(
            filename=filename, num=badge,
            title_th=ch["title_th"], title_en=ch["title_en"]
        ))

    # อัปเดต TOC
    with open(toc_path, "r", encoding="utf-8") as f:
        toc_html = f.read()
    toc_html = toc_html.replace("<!-- WORK_RULES_TOC_PLACEHOLDER -->", "\n".join(toc_items))
    with open(toc_path, "w", encoding="utf-8") as f:
        f.write(toc_html)

    print(f"\n✅ อัปเดต TOC → work-rules/index.html")
    print("\n" + "="*56)
    print("  เสร็จ! เปิด index.html ด้วย Live Server")
    print("="*56 + "\n")

if __name__ == "__main__":
    main()
