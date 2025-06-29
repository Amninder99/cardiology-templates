import os
import sys
import glob
import markdown
import pdfkit
from PyPDF2 import PdfReader, PdfWriter
from tempfile import NamedTemporaryFile

# Set the path to the Final Templates folder
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), 'Final Templates')

# Helper to convert markdown to HTML
def md_to_html(md_text, title=None):
    html = markdown.markdown(md_text, extensions=['toc', 'fenced_code'])
    if title:
        html = f'<h1>{title}</h1>\n' + html
    return f"""
    <html>
    <head>
        <meta charset='utf-8'>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 2em; }}
            h1, h2, h3, h4 {{ color: #003366; }}
            .pagenum {{ position: fixed; bottom: 10px; right: 20px; font-size: 10pt; color: #888; }}
        </style>
    </head>
    <body>{html}</body>
    </html>
    """

def add_page_numbers(input_pdf, output_pdf):
    from PyPDF2 import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from io import BytesIO

    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    num_pages = len(reader.pages)

    for i, page in enumerate(reader.pages):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont("Helvetica", 9)
        can.drawRightString(570, 10, f"Page {i+1} of {num_pages}")
        can.save()
        packet.seek(0)
        from PyPDF2 import PdfReader as RLReader
        overlay = RLReader(packet)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    with open(output_pdf, 'wb') as f:
        writer.write(f)

# Main conversion function
def convert_markdown_to_pdf(md_files, output_pdf):
    html_files = []
    bookmarks = []
    for md_file in md_files:
        with open(md_file, 'r', encoding='utf-8') as f:
            md_text = f.read()
        title = os.path.splitext(os.path.basename(md_file))[0]
        html = md_to_html(md_text, title=title)
        tmp_html = NamedTemporaryFile(delete=False, suffix='.html')
        tmp_html.write(html.encode('utf-8'))
        tmp_html.close()
        html_files.append(tmp_html.name)
        bookmarks.append(title)

    # Check if wkhtmltopdf is available
    try:
        pdfkit_config = pdfkit.configuration()
        pdfkit.from_file(html_files, output_pdf, configuration=pdfkit_config)
    except OSError as e:
        print("Error: wkhtmltopdf is not installed or not in your PATH.")
        print("Download from https://wkhtmltopdf.org/downloads.html and add to your PATH.")
        sys.exit(1)
    except Exception as e:
        print(f"PDF conversion failed: {e}")
        sys.exit(1)

    # Add page numbers
    tmp_pdf = NamedTemporaryFile(delete=False, suffix='.pdf')
    tmp_pdf.close()
    os.rename(output_pdf, tmp_pdf.name)
    add_page_numbers(tmp_pdf.name, output_pdf)
    os.unlink(tmp_pdf.name)

    # Add bookmarks (one per file, first page of each)
    reader = PdfReader(output_pdf)
    writer = PdfWriter()
    page_num = 0
    for i, md_file in enumerate(md_files):
        num_pages = 1  # Each file is at least 1 page
        writer.add_page(reader.pages[page_num])
        try:
            writer.add_outline_item(bookmarks[i], page_num)
        except Exception:
            writer.addBookmark(bookmarks[i], page_num)
        page_num += num_pages
    for i in range(page_num, len(reader.pages)):
        writer.add_page(reader.pages[i])
    with open(output_pdf, 'wb') as f:
        writer.write(f)
    for html_file in html_files:
        os.unlink(html_file)

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Convert markdown templates to PDF with bookmarks and page numbers.')
    parser.add_argument('--all', action='store_true', help='Convert all markdown files in Final Templates folder')
    parser.add_argument('--file', type=str, help='Convert a single markdown file (filename only, in Final Templates)')
    parser.add_argument('--output', type=str, default='output.pdf', help='Output PDF filename')
    args = parser.parse_args()

    if args.all:
        md_files = sorted(glob.glob(os.path.join(TEMPLATES_DIR, '*.md')))
        if not md_files:
            print('No markdown files found in Final Templates.')
            sys.exit(1)
        convert_markdown_to_pdf(md_files, args.output)
        print(f'Created PDF: {args.output}')
    elif args.file:
        md_file = os.path.join(TEMPLATES_DIR, args.file)
        if not os.path.exists(md_file):
            print(f'File not found: {md_file}')
            sys.exit(1)
        convert_markdown_to_pdf([md_file], args.output)
        print(f'Created PDF: {args.output}')
    else:
        parser.print_help()
