import os
from pathlib import Path
import win32com.client as win32

# ----------------------------------------------------------------------
# CONFIGURATION
# ----------------------------------------------------------------------
HEADER_TEXT = "23BT521 CHEMICAL ENGINEERING LABORATORY FOR BIOTECHNOLOGISTS"
FOOTER_TEXT = "GNANAMANI COLLEGE OF TECHNOLOGY"

# Word numeric constants (no win32com.client.constants used)
WD_PAPER_A4 = 7                         # WdPaperSize.wdPaperA4 :contentReference[oaicite:0]{index=0}
WD_ORIENT_PORTRAIT = 0                  # WdOrientation.wdOrientPortrait :contentReference[oaicite:1]{index=1}
WD_ALIGN_PARAGRAPH_LEFT = 0             # WdParagraphAlignment.wdAlignParagraphLeft :contentReference[oaicite:2]{index=2}
WD_ALIGN_PARAGRAPH_CENTER = 1           # WdParagraphAlignment.wdAlignParagraphCenter :contentReference[oaicite:3]{index=3}
WD_ALIGN_PARAGRAPH_RIGHT = 2            # WdParagraphAlignment.wdAlignParagraphRight :contentReference[oaicite:4]{index=4}
WD_HEADER_FOOTER_PRIMARY = 1            # WdHeaderFooterIndex.wdHeaderFooterPrimary :contentReference[oaicite:5]{index=5}
WD_FIELD_PAGE = 33                      # WdFieldType.wdFieldPage :contentReference[oaicite:6]{index=6}
WD_COLOR_GRAY25 = 12632256              # WdColor.wdColorGray25 :contentReference[oaicite:7]{index=7}
WD_COLOR_BLACK = 0                      # WdColor.wdColorBlack :contentReference[oaicite:8]{index=8}
WD_EXPORT_FORMAT_PDF = 17               # WdExportFormat.wdExportFormatPDF :contentReference[oaicite:9]{index=9}
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0        # WdExportOptimizeFor.wdExportOptimizeForPrint :contentReference[oaicite:10]{index=10}
WD_EXPORT_RANGE_ALL_DOC = 0             # WdExportRange.wdExportAllDocument :contentReference[oaicite:11]{index=11}
WD_EXPORT_ITEM_DOC_CONTENT = 0          # WdExportItem.wdExportDocumentContent :contentReference[oaicite:12]{index=12}
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1  # WdExportCreateBookmarks.wdExportCreateHeadingBookmarks :contentReference[oaicite:13]{index=13}


# ----------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------
def cm_to_points(cm: float) -> float:
    """Convert centimeters to Word points."""
    return cm * 28.3464567  # 1 cm ≈ 28.3464567 points


def set_page_setup(doc):
    """Set A4 portrait and reasonable margins."""
    ps = doc.PageSetup

    # A4 / Portrait
    ps.PaperSize = WD_PAPER_A4
    ps.Orientation = WD_ORIENT_PORTRAIT

    # Margins (in points) – adjust if you want slightly different margins
    ps.TopMargin = cm_to_points(2.5)
    ps.BottomMargin = cm_to_points(2.0)
    ps.LeftMargin = cm_to_points(2.5)
    ps.RightMargin = cm_to_points(2.0)

    # Header / footer distances
    ps.HeaderDistance = cm_to_points(1.0)
    ps.FooterDistance = cm_to_points(1.0)


def resize_images_to_fit(doc):
    """
    Scale all images so they fit within the available A4 content
    area (both horizontally and vertically).
    """
    ps = doc.PageSetup

    content_width = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    content_height = (
        ps.PageHeight
        - ps.TopMargin
        - ps.BottomMargin
        - ps.HeaderDistance
        - ps.FooterDistance
    )

    if content_width <= 0 or content_height <= 0:
        # Fallback if margins are misconfigured
        content_width = ps.PageWidth * 0.9
        content_height = ps.PageHeight * 0.8

    # InlineShapes (images in text flow)
    for ish in list(doc.InlineShapes):
        try:
            w = ish.Width
            h = ish.Height
            if w <= 0 or h <= 0:
                continue

            scale = min(1.0, content_width / w, content_height / h)
            if scale < 1.0:
                ish.Width = w * scale
                ish.Height = h * scale
        except Exception:
            # Ignore problematic shapes and continue
            continue

    # Floating shapes
    for shp in list(doc.Shapes):
        try:
            # Only resize if it has meaningful dimensions
            w = shp.Width
            h = shp.Height
            if w <= 0 or h <= 0:
                continue

            scale = min(1.0, content_width / w, content_height / h)
            if scale < 1.0:
                shp.Width = w * scale
                shp.Height = h * scale
        except Exception:
            continue


def apply_header_footer(doc):
    """
    Apply:
    - Header (right, grey, italic, small) with HEADER_TEXT
    - Footer with:
        1) Page number centered, bold, black
        2) FOOTER_TEXT right, grey, italic, small
    to every section.
    """
    for section in doc.Sections:
        # ----------------- HEADER -----------------
        header = section.Headers(WD_HEADER_FOOTER_PRIMARY)
        h_range = header.Range
        h_range.Text = HEADER_TEXT
        h_range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT

        h_font = h_range.Font
        h_font.Size = 9
        h_font.Italic = True
        h_font.Color = WD_COLOR_GRAY25

        # ----------------- FOOTER -----------------
        footer = section.Footers(WD_HEADER_FOOTER_PRIMARY)
        f_range = footer.Range

        # Clear existing footer content
        f_range.Text = ""

        # 1) Page number paragraph (centered, bold, black)
        # After setting Text = "", footer.Range still has one empty paragraph.
        page_para = footer.Range.Paragraphs(1)
        page_rng = page_para.Range

        page_rng.Text = ""  # ensure empty before inserting field
        page_rng.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER

        p_font = page_rng.Font
        p_font.Bold = True
        p_font.Italic = False
        p_font.Color = WD_COLOR_BLACK

        # Add PAGE field
        page_rng.Fields.Add(page_rng, WD_FIELD_PAGE)

        # Ensure a new paragraph after the page number
        page_rng.InsertParagraphAfter()

        # 2) Footer text paragraph (right, grey, italic, small)
        all_paras = footer.Range.Paragraphs
        footer_para = all_paras(all_paras.Count)  # last paragraph
        footer_rng = footer_para.Range

        footer_rng.Text = FOOTER_TEXT
        footer_rng.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT

        f_font = footer_rng.Font
        f_font.Size = 9
        f_font.Italic = True
        f_font.Color = WD_COLOR_GRAY25


def export_to_pdf(doc, pdf_path: Path):
    """Export the active Word document to PDF with print-quality settings."""
    doc.ExportAsFixedFormat(
        OutputFileName=str(pdf_path),
        ExportFormat=WD_EXPORT_FORMAT_PDF,
        OpenAfterExport=False,
        OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
        Range=WD_EXPORT_RANGE_ALL_DOC,
        From=1,
        To=1,
        Item=WD_EXPORT_ITEM_DOC_CONTENT,
        IncludeDocProps=True,
        KeepIRM=True,
        CreateBookmarks=WD_EXPORT_CREATE_HEADING_BOOKMARKS,
        DocStructureTags=True,
        BitmapMissingFonts=True,
        UseISO19005_1=False,
    )


def convert_all_html_to_pdf(input_dir: str, output_dir: str, visible: bool = False):
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    print(f"Input directory : {input_path}")
    print(f"Output directory: {output_path}")

    word = win32.Dispatch("Word.Application")
    word.Visible = visible

    try:
        for html_file in sorted(input_path.glob("*.html")):
            print(f"\nProcessing: {html_file.name}")
            pdf_file = output_path / (html_file.stem + ".pdf")

            try:
                doc = word.Documents.Open(str(html_file))

                # Page setup: A4 portrait, margins
                set_page_setup(doc)

                # Resize all images so they stay within A4 content area
                resize_images_to_fit(doc)

                # Apply header & footer with page numbers
                apply_header_footer(doc)

                # Export to PDF
                export_to_pdf(doc, pdf_file)

                print(f"  ✓ Saved: {pdf_file.name}")

                # Close the document without saving changes to the .docx/.html
                doc.Close(SaveChanges=False)

            except Exception as e:
                print(f"  ✖ Failed for {html_file.name}: {e}")
                try:
                    # Ensure doc is closed if partially opened
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass

    finally:
        word.Quit()


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
if __name__ == "__main__":
    # By default: process the current directory
    # You can also hard-code your HTML folder path here if you prefer.
    base_dir = os.path.dirname(os.path.abspath(__file__))
    INPUT_DIR = base_dir
    OUTPUT_DIR = base_dir

    convert_all_html_to_pdf(INPUT_DIR, OUTPUT_DIR, visible=False)
