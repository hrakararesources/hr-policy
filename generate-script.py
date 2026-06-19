#!/usr/bin/env python3
"""
========================================================
 Akara Resources — HR Policy HTML Generator
 ใช้สำหรับ: อ่านไฟล์ PDF แล้วสร้างหน้า HTML อัตโนมัติ
========================================================

วิธีใช้:
  1. วางไฟล์ PDF Work Rules ใน:  content/work-rules/
  2. วางไฟล์ PDF Welfare ใน:      content/welfare/
  3. รัน: python generate-script.py

Requirements:
  pip install pymupdf  (หรือ pip install PyMuPDF)
"""

import os
import re
import json

# ลองใช้ PyMuPDF ก่อน ถ้าไม่มีให้แจ้งวิธีติดตั้ง
try:
    import fitz  # PyMuPDF
    PDF_LIB = "pymupdf"
except ImportError:
    try:
        from pdfminer.high_level import extract_text as pdfminer_extract
        PDF_LIB = "pdfminer"
    except ImportError:
        PDF_LIB = None

# ============================================================
#  CONFIG — แก้ตรงนี้ถ้าชื่อโฟลเดอร์ต่างออกไป
# ============================================================
BASE_DIR        = os.path.dirname(os.path.abspath(__file__))
CONTENT_WR_DIR  = os.path.join(BASE_DIR, "content", "work-rules")   # โฟลเดอร์ PDF Work Rules
CONTENT_WF_DIR  = os.path.join(BASE_DIR, "content", "welfare")      # โฟลเดอร์ PDF Welfare
OUTPUT_WR_DIR   = os.path.join(BASE_DIR, "work-rules")
OUTPUT_WF_DIR   = os.path.join(BASE_DIR, "welfare")
ASSETS_PREFIX_WR = "../assets"
ASSETS_PREFIX_WF = "../assets"

# ============================================================
#  HELPERS
# ============================================================

def extract_text_from_pdf(pdf_path):
    """อ่านข้อความจาก PDF"""
    if PDF_LIB == "pymupdf":
        doc = fitz.open(pdf_path)
        pages_text = []
        for page in doc:
            pages_text.append(page.get_text("text"))
        doc.close()
        return pages_text  # list ของข้อความแต่ละหน้า
    elif PDF_LIB == "pdfminer":
        text = pdfminer_extract(pdf_path)
        return [text]
    else:
        print("❌  ไม่พบ library สำหรับอ่าน PDF")
        print("    กรุณารัน: pip install PyMuPDF")
        return []


def text_to_html_paragraphs(text):
    """แปลงข้อความดิบเป็น HTML paragraphs"""
    lines = text.strip().split("\n")
    html_parts = []
    buffer = []

    for line in lines:
        line = line.strip()
        if not line:
            if buffer:
                para = " ".join(buffer)
                # ตรวจหัวข้อ: บรรทัดสั้น + ตัวหนา หรือ ขึ้นต้นด้วยตัวเลข
                if re.match(r"^(\d+[\.\)]\s|ข้อ\s*\d+|Section\s*\d+|Article\s*\d+|หมวด|Chapter)", para, re.IGNORECASE):
                    html_parts.append(f'<h3>{para}</h3>')
                elif len(para) < 80 and not para.endswith('.'):
                    html_parts.append(f'<h4>{para}</h4>')
                else:
                    html_parts.append(f'<p>{para}</p>')
                buffer = []
        else:
            buffer.append(line)

    if buffer:
        html_parts.append(f'<p>{" ".join(buffer)}</p>')

    return "\n".join(html_parts)


def get_sorted_pdfs(folder):
    """รายชื่อไฟล์ PDF เรียงตามชื่อ"""
    if not os.path.exists(folder):
        return []
    files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
    files.sort()  # เรียงตามชื่อ (ควรตั้งชื่อ 01_xxx.pdf, 02_xxx.pdf ...)
    return files


def filename_to_title(filename):
    """แปลงชื่อไฟล์เป็นชื่อบท เช่น 01_สิทธิพนักงาน.pdf → สิทธิพนักงาน"""
    name = os.path.splitext(filename)[0]          # ตัด .pdf
    name = re.sub(r"^[\d_\-\s]+", "", name)       # ตัดเลขนำหน้า
    name = name.replace("_", " ").replace("-", " ").strip()
    return name if name else filename


# ============================================================
#  HTML TEMPLATES
# ============================================================

PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="th" data-lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{page_title} — Akara Resources</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="{assets}/style.css">
</head>
<body>

  <header class="topbar">
    <div class="topbar-left">
      <a href="index.html" class="back-btn">
        <svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
          <path d="M9 2L4 7L9 12"/>
        </svg>
        <span class="text-th">สารบัญ</span>
        <span class="text-en">Contents</span>
      </a>
      <img src="{assets}/Akara_Logo.jpg" alt="Akara Resources" class="topbar-logo">
      <div class="topbar-title">
        <span class="text-th">{section_th}</span>
        <span class="text-en">{section_en}</span>
        <span>{badge_label} {num}</span>
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
      ☰ <span class="text-th">เนื้อหา</span><span class="text-en">Sections</span>
    </button>

    <div class="chapter-layout">
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-title text-th">เนื้อหาในบทนี้</div>
        <div class="sidebar-title text-en">In this chapter</div>
        <nav class="sidebar-nav" id="sidebarNav">
          <!-- จะถูก populate โดย JS -->
        </nav>
      </aside>

      <article class="chapter-card">
        <div class="chapter-badge">{badge_label} {num}</div>
        <h1 class="chapter-title">{title}</h1>
        <div class="chapter-divider"></div>
        <div class="chapter-body" id="chapterBody">
{content_html}
        </div>

        <nav class="chapter-nav">
          {prev_link}
          {next_link}
        </nav>
      </article>
    </div>
  </main>

  <footer class="footer">
    © 2026 Akara Resources Co., Ltd. — Human Resources Department
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
    const headings = document.querySelectorAll('.chapter-body h3, .chapter-body h4');
    const nav = document.getElementById('sidebarNav');
    headings.forEach((h, i) => {{
      h.id = 'section-' + i;
      const a = document.createElement('a');
      a.href = '#section-' + i;
      a.textContent = h.textContent;
      if (h.tagName === 'H4') a.style.paddingLeft = '18px';
      nav.appendChild(a);
    }});

    // Highlight active sidebar item on scroll
    const observer = new IntersectionObserver(entries => {{
      entries.forEach(e => {{
        if (e.isIntersecting) {{
          nav.querySelectorAll('a').forEach(a => a.classList.remove('active'));
          const active = nav.querySelector(`a[href="#${{e.target.id}}"]`);
          if (active) active.classList.add('active');
        }}
      }});
    }}, {{ rootMargin: '-20% 0px -75% 0px' }});
    headings.forEach(h => observer.observe(h));

    function toggleSidebar() {{
      document.getElementById('sidebar').classList.toggle('open');
    }}
  </script>
</body>
</html>
"""

TOC_ITEM_TEMPLATE = """<a href="{filename}" class="toc-item">
  <div class="toc-num{num_class}">{num:02d}</div>
  <div class="toc-text">
    <strong>{title}</strong>
    <span>{filename}</span>
  </div>
  <span class="toc-chevron">›</span>
</a>"""


# ============================================================
#  GENERATOR FUNCTIONS
# ============================================================

def generate_section(
    pdf_folder,
    output_folder,
    assets_prefix,
    section_th,
    section_en,
    badge_label,
    file_prefix,
    toc_placeholder,
    toc_index_path,
    num_class=""
):
    pdfs = get_sorted_pdfs(pdf_folder)

    if not pdfs:
        print(f"⚠️  ไม่พบไฟล์ PDF ใน {pdf_folder}")
        print(f"    วางไฟล์ PDF ของคุณแล้วรัน script อีกครั้ง")
        return

    print(f"\n📂 พบ {len(pdfs)} ไฟล์ใน {pdf_folder}")
    toc_items_html = []
    total = len(pdfs)

    for idx, pdf_file in enumerate(pdfs, start=1):
        pdf_path = os.path.join(pdf_folder, pdf_file)
        title    = filename_to_title(pdf_file)
        out_name = f"{file_prefix}-{idx:02d}.html"
        out_path = os.path.join(output_folder, out_name)

        print(f"  [{idx:02d}/{total}] {pdf_file} → {out_name}")

        # อ่าน PDF
        pages = extract_text_from_pdf(pdf_path)
        all_text = "\n\n".join(pages)
        content_html = text_to_html_paragraphs(all_text)

        if not content_html.strip():
            content_html = f'<p style="color:var(--text-muted);font-style:italic;">ไม่สามารถแยกข้อความจาก PDF นี้ได้อัตโนมัติ กรุณาเปิดไฟล์ {pdf_file} แล้ว copy เนื้อหามาวางที่นี่</p>'

        # Prev / Next links
        prev_link = ""
        next_link = ""
        if idx > 1:
            prev_file = f"{file_prefix}-{idx-1:02d}.html"
            prev_title = filename_to_title(pdfs[idx-2])
            prev_link = f'<a href="{prev_file}">← {prev_title[:35]}{"..." if len(prev_title)>35 else ""}</a>'
        if idx < total:
            next_file = f"{file_prefix}-{idx+1:02d}.html"
            next_title = filename_to_title(pdfs[idx])
            next_link = f'<a href="{next_file}" class="next">{next_title[:35]}{"..." if len(next_title)>35 else ""} →</a>'

        # Write HTML
        html = PAGE_TEMPLATE.format(
            page_title   = title,
            assets       = assets_prefix,
            section_th   = section_th,
            section_en   = section_en,
            badge_label  = badge_label,
            num          = idx,
            title        = title,
            content_html = content_html,
            prev_link    = prev_link,
            next_link    = next_link,
        )

        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)

        # TOC item
        toc_items_html.append(TOC_ITEM_TEMPLATE.format(
            filename  = out_name,
            num       = idx,
            num_class = num_class,
            title     = title,
        ))

    # Inject TOC into index.html
    with open(toc_index_path, "r", encoding="utf-8") as f:
        toc_html = f.read()

    toc_block = "\n".join(toc_items_html)
    toc_html  = toc_html.replace(f"<!-- {toc_placeholder} -->", toc_block)

    with open(toc_index_path, "w", encoding="utf-8") as f:
        f.write(toc_html)

    print(f"  ✅ อัปเดต TOC → {toc_index_path}")
    print(f"  ✅ สร้างไฟล์ HTML {total} ไฟล์ เรียบร้อย")


# ============================================================
#  MAIN
# ============================================================

def main():
    print("=" * 54)
    print("  Akara Resources — HR Policy HTML Generator")
    print("=" * 54)

    if PDF_LIB is None:
        print("\n❌  กรุณาติดตั้ง PyMuPDF ก่อน:")
        print("    pip install PyMuPDF\n")
        return

    print(f"\n✔  ใช้ library: {PDF_LIB}")

    # Work Rules
    generate_section(
        pdf_folder     = CONTENT_WR_DIR,
        output_folder  = OUTPUT_WR_DIR,
        assets_prefix  = ASSETS_PREFIX_WR,
        section_th     = "ระเบียบข้อบังคับการทำงาน",
        section_en     = "Work Rules & Regulations",
        badge_label    = "บทที่",
        file_prefix    = "chapter",
        toc_placeholder= "WORK_RULES_TOC_PLACEHOLDER",
        toc_index_path = os.path.join(OUTPUT_WR_DIR, "index.html"),
        num_class      = "",
    )

    # Welfare
    generate_section(
        pdf_folder     = CONTENT_WF_DIR,
        output_folder  = OUTPUT_WF_DIR,
        assets_prefix  = ASSETS_PREFIX_WF,
        section_th     = "สวัสดิการและผลประโยชน์",
        section_en     = "Welfare & Benefits",
        badge_label    = "หัวข้อที่",
        file_prefix    = "welfare",
        toc_placeholder= "WELFARE_TOC_PLACEHOLDER",
        toc_index_path = os.path.join(OUTPUT_WF_DIR, "index.html"),
        num_class      = " gold-bg",
    )

    print("\n" + "=" * 54)
    print("  ✅ เสร็จสมบูรณ์! เปิด index.html ด้วย Live Server")
    print("=" * 54 + "\n")


if __name__ == "__main__":
    main()
