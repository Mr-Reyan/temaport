import os
import re
import openpyxl

# Configuration Paths
REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
PORTS_DIR = os.path.join(REPO_ROOT, "ports")
CSS_FILE = os.path.join(REPO_ROOT, "assets", "css", "port-pages.css")
EXCEL_PATH = r"D:\FalconXoft Ship Chandler\Website Data\Ship_Chandler_SEO_49_Pages_FULL.xlsx"

# Section heading IDs matching Table of Contents and Bissau layout
H2_IDS = [
    'professional-services',
    'provisions',
    'bonded',
    'deck',
    'engine',
    'cabin',
    'safety',
    'electrical',
    'chemicals',
    'welding',
    'spare-parts',
    'publications',
    'tools',
    'technical-support',
    'urgent',
    'vessels',
    'rfq',
    'why-choose',
    None,
    None,
    None,
    None,
]

def extract_and_centralize_css(bissau_html):
    """
    Extracts the inline CSS from bissau-ship-chandler-guinea.html and saves
    it to assets/css/port-pages.css, then replaces the <style> block with
    <link rel="stylesheet" href="../assets/css/port-pages.css">.
    """
    style_start = '<style>'
    style_end = '</style>'
    
    if style_start in bissau_html and style_end in bissau_html:
        p1 = bissau_html.find(style_start)
        p2 = bissau_html.find(style_end) + len(style_end)
        
        css_content = bissau_html[p1 + len(style_start) : bissau_html.find(style_end)].strip()
        
        # Ensure parent directory exists and write CSS
        os.makedirs(os.path.dirname(CSS_FILE), exist_ok=True)
        with open(CSS_FILE, 'w', encoding='utf-8') as f:
            f.write(css_content + '\n')
            
        print(f"[CSS] Extracted {len(css_content)} bytes of styles into {CSS_FILE}")
        
        # Replace inline <style> block with link tag
        replacement = '  <!-- Centralized Port Pages CSS -->\n  <link rel="stylesheet" href="../assets/css/port-pages.css">'
        bissau_html = bissau_html[:p1] + replacement + bissau_html[p2:]
    else:
        print("[CSS] Master template already links to external CSS.")
        
    return bissau_html


def render_article_html(desc_text):
    """
    Parses the Markdown description from Excel into semantic HTML
    matching the exact structure and indentation of bissau-ship-chandler-guinea.html
    """
    blocks = [b.strip() for b in desc_text.split('\n\n') if b.strip()]
    lines = []
    h2_index = 0
    
    for b in blocks:
        if b.startswith('# '):
            # Skip top-level H1 (rendered in Hero section)
            continue
        elif b.startswith('## '):
            heading_text = b[3:].strip()
            hid = H2_IDS[h2_index] if h2_index < len(H2_IDS) else None
            h2_index += 1
            if lines:
                lines.append('')
            if hid:
                lines.append(f'          <h2 id="{hid}">{heading_text}</h2>')
            else:
                lines.append(f'          <h2>{heading_text}</h2>')
        elif b.startswith('- '):
            lines.append('          <ul>')
            for item in b.split('\n'):
                item = item.strip()
                if not item:
                    continue
                val = re.sub(r'^\s*-\s*', '', item)
                lines.append(f'            <li>{val}</li>')
            lines.append('          </ul>')
        else:
            # Paragraph or bold summary
            if b.startswith('**') and b.endswith('**'):
                bold_content = b[2:-2].strip()
                lines.append(f'          <p><strong>{bold_content}</strong></p>')
            else:
                clean_p = ' '.join(b.split())
                lines.append(f'          <p>{clean_p}</p>')
                
    return '\n'.join(lines)


def get_template_parts():
    """
    Extracts the master HTML shell from bissau-ship-chandler-guinea.html,
    centralizing the CSS into port-pages.css in the process.
    """
    bissau_path = os.path.join(PORTS_DIR, "bissau-ship-chandler-guinea.html")
    with open(bissau_path, 'r', encoding='utf-8') as f:
        master_html = f.read()

    master_html = extract_and_centralize_css(master_html)

    # Use local assets for logo and hero background
    master_html = master_html.replace('https://i.ibb.co/6803z6K/logo.png', '../assets/imgs/logo.png')
    master_html = master_html.replace('https://images.unsplash.com/photo-1494412574643-ff11b0a5c1c3?q=80&w=2070&auto=format&fit=crop', '../assets/imgs/bg-hero.jpg')

    art_start = '<article class="article-card">\n'
    art_end = '\n        </article>'
    p1 = master_html.find(art_start) + len(art_start)
    p2 = master_html.find(art_end, p1)

    if p1 == -1 or p2 == -1:
        raise ValueError("Could not find article markers in bissau-ship-chandler-guinea.html")

    head_part = master_html[:p1]
    tail_part = master_html[p2:]
    return head_part, tail_part


def generate_page_html(head_part, tail_part, page_name, h1, meta_title, meta_desc, full_desc):
    """
    Generates complete HTML for a port page matching Bissau structure
    """
    orig_name = "Bissau Ship Chandler - Guinea"
    orig_title = "Bissau Ship Chandler - Guinea | Ship Supply, Provisions & Marine Stores"
    orig_desc = "Bissau Ship Chandler - Guinea provides professional ship supply, provisions, bonded stores, deck and engine stores, safety equipment, spare parts and marine procurement services."
    
    h = head_part
    # Replace Title
    h = h.replace(f"<title>{orig_title}</title>", f"<title>{meta_title}</title>")
    # Replace Meta Description and OG Description (orig_desc appears in both)
    h = h.replace(f'content="{orig_desc}"', f'content="{meta_desc}"')
    # Replace OG Title
    h = h.replace(f'<meta property="og:title" content="{orig_title}">', f'<meta property="og:title" content="{meta_title}">')
    # Replace Hero image alt
    h = h.replace(f'alt="{orig_name}"', f'alt="{h1}"')
    # Replace Breadcrumb
    h = h.replace(f'<li>{orig_name}</li>', f'<li>{page_name}</li>')
    # Replace H1
    h = h.replace(f'<h1 class="hero-title">{orig_name}</h1>', f'<h1 class="hero-title">{h1}</h1>')
    
    # Render article content from Excel description
    art = render_article_html(full_desc)
    
    return h + art + tail_part


def main():
    print(f"Reading Excel data from: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet = wb['SEO Pages']
    
    head_part, tail_part = get_template_parts()
    
    rows = list(sheet.iter_rows(values_only=True))[1:]  # Skip header row
    print(f"Found {len(rows)} port rows in Excel.")
    
    generated_count = 0
    generated_files = []
    
    for row in rows:
        num, page_name, h1, meta_title, meta_desc, full_desc = row[:6]
        if not page_name:
            continue
            
        slug = re.sub(r'[^a-z0-9]+', '-', str(page_name).lower()).strip('-') + '.html'
        out_path = os.path.join(PORTS_DIR, slug)
        
        html_content = generate_page_html(
            head_part=head_part,
            tail_part=tail_part,
            page_name=page_name,
            h1=h1,
            meta_title=meta_title,
            meta_desc=meta_desc,
            full_desc=full_desc
        )
        
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
            
        generated_count += 1
        generated_files.append((num, page_name, slug))
        print(f"[{num:02d}/49] Generated: {slug} ({page_name})")
        
    print("\n" + "="*60)
    print(f"SUCCESS: Processed {generated_count} rows across {len(set(f[2] for f in generated_files))} unique port page files.")
    print(f"Centralized CSS written to: {CSS_FILE}")
    print("="*60)

if __name__ == "__main__":
    main()
