import os
from pathlib import Path

import win32com.client as win32
from win32com.client import constants


def set_page_setup(doc):
    """Set A4, portrait, and (optionally) margins and header/footer behaviour."""
    ps = doc.PageSetup

    # A4, portrait
    ps.PaperSize = constants.wdPaperA4
    ps.Orientation = constants.wdOrientPortrait

    # Ensure same header/footer on all pages
    ps.DifferentFirstPageHeaderFooter = False
    ps.OddAndEvenPagesHeaderFooter = False

    # OPTIONAL: adjust margins here if you want (values in points; 1 inch = 72 pt)
    # For example, ~2.5 cm ≈ 71 pt
    # ps.TopMargin = 71
    # ps.BottomMargin = 71
    # ps.LeftMargin = 71
    # ps.RightMargin = 71


def apply_header_footer(doc):
    """
    Apply uniform header and footer on all sections:

      HEADER  (right aligned, grey, italics, small font):
        23BT521 CHEMICAL ENGINEERING LABORATORY FOR BIOTECHNOLOGISTS

      FOOTER  (right aligned, grey, italics, small font):
        GNANAMANI COLLEGE OF TECHNOLOGY

      PAGE NUMBER (separate line, centered, bold, black, normal font):
        dynamic page number starting from 1
    """
    header_text = "23BT521 CHEMICAL ENGINEERING LABORATORY FOR BIOTECHNOLOGISTS"
    footer_text = "GNANAMANI COLLEGE OF TECHNOLOGY"

    for idx, section in enumerate(doc.Sections, start=1):
        # ------------------------
        # HEADER: right aligned
        # ------------------------
        header_range = section.Headers(constants.wdHeaderFooterPrimary).Range
        header_range.Text = header_text
        header_range.ParagraphFormat.Alignment = constants.wdAlignParagraphRight

        header_font = header_range.Font
        header_font.Size = 9          # "small" font size
        header_font.Italic = True
        header_font.Color = constants.wdColorGrayText

        # ------------------------
        # FOOTER: right text + centered page number
        # ------------------------
        footer_obj = section.Footers(constants.wdHeaderFooterPrimary)
        footer_range = footer_obj.Range

        # Clear existing footer content
        footer_range.Text = ""
        footer_range.ParagraphFormat.Alignment = constants.wdAlignParagraphLeft  # reset

        # 1) Insert college name (right aligned, grey, italics, small)
        footer_range.Text = footer_text
        footer_range.ParagraphFormat.Alignment = constants.wdAlignParagraphRight

        footer_font = footer_range.Font
        footer_font.Size = 9
        footer_font.Italic = True
        footer_font.Color = constants.wdColorGrayText
        footer_font.Bold = False

        # Move to end to add a new paragraph for the page number
        footer_range.Collapse(constants.wdCollapseEnd)
        footer_range.InsertParagraphAfter()
        footer_range.Collapse(constants.wdCollapseEnd)

        # 2) Page number paragraph: centered, bold, black, normal (non-italic)
        page_range = footer_range
        page_range.ParagraphFormat.Alignment = constants.wdAlignParagraphCenter

        # Add PAGE field for dynamic numbering
        field = page_range.Fields.Add(
            Range=page_range,
            Type=constants.wdFieldPage
        )

        page_font = page_range.Font
        page_font.Size = 10          # slightly larger if you wish
        page_font.Bold = True
        page_font.Italic = False
        page_font.Color = constants.wdColorAutomatic  # black

        # Ensure page numbering starts at 1 for the first section
        # and continues (or restarts) consistently.
        pn = footer_obj.PageNumbers
        if idx == 1:
            pn.RestartNumberingAtSection = True
            pn.StartingNumber = 1
        else:
            # Continue numbering by default in later sections
            pn.RestartNumberingAtSection = False


def resize_images_to_fit(doc):
    """
    For all images (InlineShapes and Shapes), ensure they fit
    inside the printable area (page size minus margins).
    If an image is too large, scale it down while preserving aspect ratio.
    """
    ps = doc.PageSetup
    page_width = ps.PageWidth       # in points
    page_height = ps.PageHeight     # in points

    max_width = page_width - ps.LeftMargin - ps.RightMargin
    max_height = page_height - ps.TopMargin - ps.BottomMargin

    # Safety: avoid negative values in case of strange settings
    max_width = max(max_width, 1)
    max_height = max(max_height, 1)

    def scale_shape(shape):
        """Scale a single shape-like object if it is too large."""
        try:
            w = shape.Width
            h = shape.Height
        except Exception:
            return  # skip if shape has no dimension

        if w <= 0 or h <= 0:
            return

        # Determine scale factor so that both width and height fit
        scale_w = max_width / w
        scale_h = max_height / h
        scale = min(1.0, scale_w, scale_h)  # only scale down (never enlarge)

        if scale < 1.0:
            try:
                shape.LockAspectRatio = False
            except Exception:
                pass

            shape.Width = w * scale
            shape.Height = h * scale

    # Inline images
    for ishape in doc.InlineShapes:
        scale_shape(ishape)

    # Floating shapes
    for shp in doc.Shapes:
        scale_shape(shp)


def convert_all_html_to_pdf(input_dir=".", output_dir=None, visible=False):
    """
    Converts all .html and .htm files in `input_dir` to PDF using Microsoft Word.
    - Sets A4 page size and portrait orientation.
    - Rescales images to fit inside page margins.
    - Adds required header and footer (right aligned).
    - Adds dynamic page numbers in footer (centered, bold, black, starting at 1).
    """
    input_path = Path(input_dir).resolve()
    if output_dir is None:
        output_path = input_path
    else:
        output_path = Path(output_dir).resolve()
        output_path.mkdir(parents=True, exist_ok=True)

    print(f"Input directory : {input_path}")
    print(f"Output directory: {output_path}")

    # Collect HTML files
    html_files = sorted(
        list(input_path.glob("*.html")) + list(input_path.glob("*.htm"))
    )

    if not html_files:
        print("No HTML files found. Nothing to convert.")
        return

    # Start Word
    word = win32.Dispatch("Word.Application")
    word.Visible = visible

    wdFormatPDF = 17

    try:
        for html_file in html_files:
            print(f"\nProcessing: {html_file.name}")
            pdf_file = output_path / (html_file.stem + ".pdf")

            # Open HTML
            doc = word.Documents.Open(str(html_file))

            try:
                # Page setup
                set_page_setup(doc)

                # Header, footer, and page number
                apply_header_footer(doc)

                # Resize images
                resize_images_to_fit(doc)

                # Export to PDF
                print(f"  -> Exporting to: {pdf_file.name}")
                doc.ExportAsFixedFormat(
                    OutputFileName=str(pdf_file),
                    ExportFormat=wdFormatPDF,
                    OpenAfterExport=False,
                    OptimizeFor=constants.wdExportOptimizeForPrint,
                    Item=constants.wdExportDocumentContent,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=constants.wdExportCreateHeadingBookmarks,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,  # set True for PDF/A
                )
                print("  ✔ Done")

            except Exception as e:
                print(f"  ✖ Failed for {html_file.name}: {e}")

            finally:
                doc.Close(False)

    finally:
        word.Quit()

    print("\nAll conversions finished.")


if __name__ == "__main__":
    # Adjust these paths if needed
    INPUT_DIR = "."        # folder containing your HTML files
    OUTPUT_DIR = None      # None = same folder; or e.g. r".\pdf_output"

    convert_all_html_to_pdf(INPUT_DIR, OUTPUT_DIR)
