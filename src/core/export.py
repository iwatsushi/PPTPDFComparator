"""Export functionality for comparison reports."""

from __future__ import annotations

import base64
import io
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Optional

from PIL import Image

if TYPE_CHECKING:
    try:
        from src.core.document import Document
        from src.core.page_matcher import MatchingResult, MatchResult
        from src.core.exclusion_zone import ExclusionZoneSet
    except ImportError:
        from core.document import Document
        from core.page_matcher import MatchingResult, MatchResult
        from core.exclusion_zone import ExclusionZoneSet


@dataclass
class ExportConfig:
    """Configuration for export."""

    include_identical: bool = False  # Include pages with no differences
    include_thumbnails: bool = True  # Include thumbnail images
    thumbnail_width: int = 400  # Thumbnail width in pixels
    title: str = "Document Comparison Report"
    show_exclusion_zones: bool = True


def _pil_to_base64(image: Image.Image, format: str = "PNG") -> str:
    """Convert PIL Image to base64 string."""
    buffer = io.BytesIO()
    image.save(buffer, format=format)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def _resize_image(image: Image.Image, max_width: int) -> Image.Image:
    """Resize image maintaining aspect ratio."""
    if image.width <= max_width:
        return image
    ratio = max_width / image.width
    new_height = int(image.height * ratio)
    return image.resize((max_width, new_height), Image.Resampling.LANCZOS)


def export_to_html(
    left_doc: 'Document',
    right_doc: 'Document',
    matching_result: 'MatchingResult',
    diff_scores: dict,
    output_path: str | Path,
    config: Optional[ExportConfig] = None,
    exclusion_zones: Optional['ExclusionZoneSet'] = None,
) -> None:
    """Export comparison report to HTML.

    Args:
        left_doc: Left document
        right_doc: Right document
        matching_result: Matching results
        diff_scores: Dictionary of (left_idx, right_idx) -> has_differences
        output_path: Output HTML file path
        config: Export configuration
        exclusion_zones: Exclusion zones to display
    """
    try:
        from src.core.page_matcher import MatchStatus
    except ImportError:
        from core.page_matcher import MatchStatus

    if config is None:
        config = ExportConfig()

    output_path = Path(output_path)

    # Gather data
    matched_with_diff = []
    matched_identical = []
    unmatched_left = []
    unmatched_right = []

    for match in matching_result.matches:
        if match.status == MatchStatus.MATCHED:
            has_diff = diff_scores.get((match.left_index, match.right_index), False)
            if has_diff:
                matched_with_diff.append(match)
            else:
                matched_identical.append(match)
        elif match.status == MatchStatus.UNMATCHED_LEFT:
            unmatched_left.append(match)
        elif match.status == MatchStatus.UNMATCHED_RIGHT:
            unmatched_right.append(match)

    # Generate HTML
    html_parts = [
        '<!DOCTYPE html>',
        '<html lang="ja">',
        '<head>',
        '<meta charset="UTF-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
        f'<title>{config.title}</title>',
        '<style>',
        _get_html_styles(),
        '</style>',
        '</head>',
        '<body>',
        '<div class="container">',
        f'<h1>{config.title}</h1>',
        '<div class="meta">',
        f'<p>Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>',
        f'<p>Left Document: <strong>{left_doc.name}</strong> ({left_doc.page_count} pages)</p>',
        f'<p>Right Document: <strong>{right_doc.name}</strong> ({right_doc.page_count} pages)</p>',
        '</div>',
        '<div class="summary">',
        f'<div class="stat diff">Pages with Differences: <strong>{len(matched_with_diff)}</strong></div>',
        f'<div class="stat identical">Identical Pages: <strong>{len(matched_identical)}</strong></div>',
        f'<div class="stat unmatched">Unmatched Left: <strong>{len(unmatched_left)}</strong></div>',
        f'<div class="stat unmatched">Unmatched Right: <strong>{len(unmatched_right)}</strong></div>',
        '</div>',
    ]

    # Pages with differences
    if matched_with_diff:
        html_parts.append('<h2 class="section-diff">Pages with Differences</h2>')
        html_parts.append('<div class="pages">')
        for match in matched_with_diff:
            html_parts.append(_render_match_html(
                match, left_doc, right_doc, config, has_diff=True
            ))
        html_parts.append('</div>')

    # Identical pages (optional)
    if config.include_identical and matched_identical:
        html_parts.append('<h2 class="section-identical">Identical Pages</h2>')
        html_parts.append('<div class="pages">')
        for match in matched_identical:
            html_parts.append(_render_match_html(
                match, left_doc, right_doc, config, has_diff=False
            ))
        html_parts.append('</div>')

    # Unmatched pages
    if unmatched_left:
        html_parts.append('<h2 class="section-unmatched">Unmatched Pages (Left Only)</h2>')
        html_parts.append('<div class="pages">')
        for match in unmatched_left:
            html_parts.append(_render_unmatched_html(
                match.left_index, left_doc, "left", config
            ))
        html_parts.append('</div>')

    if unmatched_right:
        html_parts.append('<h2 class="section-unmatched">Unmatched Pages (Right Only)</h2>')
        html_parts.append('<div class="pages">')
        for match in unmatched_right:
            html_parts.append(_render_unmatched_html(
                match.right_index, right_doc, "right", config
            ))
        html_parts.append('</div>')

    html_parts.extend([
        '</div>',
        '</body>',
        '</html>',
    ])

    # Write output
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html_parts))


def _get_html_styles() -> str:
    """Get CSS styles for HTML export."""
    return '''
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: #f5f5f5;
            color: #333;
            line-height: 1.6;
        }
        .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
        h1 { color: #2c3e50; margin-bottom: 20px; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        h2 { margin: 30px 0 15px; padding: 10px; border-radius: 5px; }
        .section-diff { background: #fee; color: #c0392b; }
        .section-identical { background: #efe; color: #27ae60; }
        .section-unmatched { background: #fec; color: #e67e22; }
        .meta { background: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .meta p { margin: 5px 0; }
        .summary { display: flex; gap: 15px; flex-wrap: wrap; margin-bottom: 30px; }
        .stat { padding: 15px 25px; border-radius: 8px; color: white; }
        .stat.diff { background: #e74c3c; }
        .stat.identical { background: #27ae60; }
        .stat.unmatched { background: #f39c12; }
        .pages { display: flex; flex-direction: column; gap: 20px; }
        .page-pair {
            background: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .page-pair.has-diff { border-left: 5px solid #e74c3c; }
        .page-pair.identical { border-left: 5px solid #27ae60; }
        .page-pair.unmatched { border-left: 5px solid #f39c12; }
        .page-header {
            display: flex;
            justify-content: space-between;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 1px solid #eee;
        }
        .page-label { font-weight: bold; color: #666; }
        .page-images { display: flex; gap: 20px; justify-content: center; flex-wrap: wrap; }
        .page-image { text-align: center; }
        .page-image img {
            max-width: 100%;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .page-image p { margin-top: 8px; font-size: 14px; color: #666; }
        @media (max-width: 900px) {
            .page-images { flex-direction: column; align-items: center; }
        }
    '''


def _render_match_html(
    match: 'MatchResult',
    left_doc: 'Document',
    right_doc: 'Document',
    config: ExportConfig,
    has_diff: bool
) -> str:
    """Render HTML for a matched page pair."""
    left_page = left_doc.pages[match.left_index]
    right_page = right_doc.pages[match.right_index]

    diff_class = "has-diff" if has_diff else "identical"
    status_text = "Differences detected" if has_diff else "Identical"

    parts = [
        f'<div class="page-pair {diff_class}">',
        '<div class="page-header">',
        f'<span class="page-label">Left Page {match.left_index + 1} ↔ Right Page {match.right_index + 1}</span>',
        f'<span class="status">{status_text}</span>',
        '</div>',
        '<div class="page-images">',
    ]

    if config.include_thumbnails:
        # Left image
        if left_page.thumbnail:
            resized = _resize_image(left_page.thumbnail, config.thumbnail_width)
            b64 = _pil_to_base64(resized)
            parts.append(f'<div class="page-image"><img src="data:image/png;base64,{b64}" alt="Left Page {match.left_index + 1}"><p>Left: Page {match.left_index + 1}</p></div>')

        # Right image
        if right_page.thumbnail:
            resized = _resize_image(right_page.thumbnail, config.thumbnail_width)
            b64 = _pil_to_base64(resized)
            parts.append(f'<div class="page-image"><img src="data:image/png;base64,{b64}" alt="Right Page {match.right_index + 1}"><p>Right: Page {match.right_index + 1}</p></div>')

    parts.extend(['</div>', '</div>'])
    return '\n'.join(parts)


def _render_unmatched_html(
    page_index: int,
    doc: 'Document',
    side: str,
    config: ExportConfig
) -> str:
    """Render HTML for an unmatched page."""
    page = doc.pages[page_index]
    side_label = "Left" if side == "left" else "Right"

    parts = [
        '<div class="page-pair unmatched">',
        '<div class="page-header">',
        f'<span class="page-label">{side_label} Page {page_index + 1} (No Match)</span>',
        '</div>',
        '<div class="page-images">',
    ]

    if config.include_thumbnails and page.thumbnail:
        resized = _resize_image(page.thumbnail, config.thumbnail_width)
        b64 = _pil_to_base64(resized)
        parts.append(f'<div class="page-image"><img src="data:image/png;base64,{b64}" alt="{side_label} Page {page_index + 1}"><p>{side_label}: Page {page_index + 1}</p></div>')

    parts.extend(['</div>', '</div>'])
    return '\n'.join(parts)


def export_to_pdf(
    left_doc: 'Document',
    right_doc: 'Document',
    matching_result: 'MatchingResult',
    diff_scores: dict,
    output_path: str | Path,
    config: Optional[ExportConfig] = None,
    exclusion_zones: Optional['ExclusionZoneSet'] = None,
) -> None:
    """Export comparison report to PDF.

    Args:
        left_doc: Left document
        right_doc: Right document
        matching_result: Matching results
        diff_scores: Dictionary of (left_idx, right_idx) -> has_differences
        output_path: Output PDF file path
        config: Export configuration
        exclusion_zones: Exclusion zones to display
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        Image as RLImage, PageBreak
    )
    try:
        from src.core.page_matcher import MatchStatus
    except ImportError:
        from core.page_matcher import MatchStatus

    if config is None:
        config = ExportConfig()

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Create document
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=landscape(A4),
        rightMargin=15*mm,
        leftMargin=15*mm,
        topMargin=15*mm,
        bottomMargin=15*mm,
    )

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=20,
    )
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=16,
        spaceBefore=20,
        spaceAfter=10,
    )
    normal_style = styles['Normal']

    # Build content
    story = []

    # Title
    story.append(Paragraph(config.title, title_style))
    story.append(Spacer(1, 10))

    # Metadata
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", normal_style))
    story.append(Paragraph(f"Left Document: <b>{left_doc.name}</b> ({left_doc.page_count} pages)", normal_style))
    story.append(Paragraph(f"Right Document: <b>{right_doc.name}</b> ({right_doc.page_count} pages)", normal_style))
    story.append(Spacer(1, 20))

    # Gather data
    matched_with_diff = []
    matched_identical = []
    unmatched_left = []
    unmatched_right = []

    for match in matching_result.matches:
        if match.status == MatchStatus.MATCHED:
            has_diff = diff_scores.get((match.left_index, match.right_index), False)
            if has_diff:
                matched_with_diff.append(match)
            else:
                matched_identical.append(match)
        elif match.status == MatchStatus.UNMATCHED_LEFT:
            unmatched_left.append(match)
        elif match.status == MatchStatus.UNMATCHED_RIGHT:
            unmatched_right.append(match)

    # Summary table
    summary_data = [
        ['Category', 'Count'],
        ['Pages with Differences', str(len(matched_with_diff))],
        ['Identical Pages', str(len(matched_identical))],
        ['Unmatched (Left)', str(len(unmatched_left))],
        ['Unmatched (Right)', str(len(unmatched_right))],
    ]

    summary_table = Table(summary_data, colWidths=[150, 80])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#ffeeee')),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#eeffee')),
        ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor('#fff5e6')),
        ('BACKGROUND', (0, 4), (-1, 4), colors.HexColor('#fff5e6')),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#cccccc')),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 30))

    # Pages with differences
    if matched_with_diff:
        story.append(Paragraph("Pages with Differences", heading_style))

        for match in matched_with_diff:
            story.append(_render_match_pdf(
                match, left_doc, right_doc, config, has_diff=True
            ))
            story.append(Spacer(1, 15))

    # Identical pages (optional)
    if config.include_identical and matched_identical:
        story.append(PageBreak())
        story.append(Paragraph("Identical Pages", heading_style))

        for match in matched_identical:
            story.append(_render_match_pdf(
                match, left_doc, right_doc, config, has_diff=False
            ))
            story.append(Spacer(1, 15))

    # Unmatched pages
    if unmatched_left or unmatched_right:
        story.append(PageBreak())
        story.append(Paragraph("Unmatched Pages", heading_style))

        if unmatched_left:
            story.append(Paragraph("Left Document Only:", normal_style))
            for match in unmatched_left:
                story.append(_render_unmatched_pdf(
                    match.left_index, left_doc, "left", config
                ))
                story.append(Spacer(1, 10))

        if unmatched_right:
            story.append(Paragraph("Right Document Only:", normal_style))
            for match in unmatched_right:
                story.append(_render_unmatched_pdf(
                    match.right_index, right_doc, "right", config
                ))
                story.append(Spacer(1, 10))

    # Build PDF
    doc.build(story)


def _render_match_pdf(
    match: 'MatchResult',
    left_doc: 'Document',
    right_doc: 'Document',
    config: ExportConfig,
    has_diff: bool,
):
    """Render PDF element for a matched page pair."""
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import Table, TableStyle, Paragraph
    from reportlab.lib.styles import getSampleStyleSheet

    styles = getSampleStyleSheet()

    left_page = left_doc.pages[match.left_index]
    right_page = right_doc.pages[match.right_index]

    status_text = "DIFF" if has_diff else "OK"
    status_color = colors.HexColor('#e74c3c') if has_diff else colors.HexColor('#27ae60')

    data = [[
        Paragraph(f"Left Page {match.left_index + 1}", styles['Normal']),
        Paragraph(f"<font color='{'red' if has_diff else 'green'}'>{status_text}</font>", styles['Normal']),
        Paragraph(f"Right Page {match.right_index + 1}", styles['Normal']),
    ]]

    # Add images if available
    if config.include_thumbnails:
        left_img = _pil_to_reportlab_image(left_page.thumbnail, 120*mm) if left_page.thumbnail else ""
        right_img = _pil_to_reportlab_image(right_page.thumbnail, 120*mm) if right_page.thumbnail else ""
        data.append([left_img, "", right_img])

    table = Table(data, colWidths=[130*mm, 20*mm, 130*mm])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (1, 0), (1, 0), status_color),
        ('TEXTCOLOR', (1, 0), (1, 0), colors.white),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#dddddd')),
    ]))

    return table


def _render_unmatched_pdf(
    page_index: int,
    doc: 'Document',
    side: str,
    config: ExportConfig,
):
    """Render PDF element for an unmatched page."""
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import Table, TableStyle, Paragraph
    from reportlab.lib.styles import getSampleStyleSheet

    styles = getSampleStyleSheet()
    page = doc.pages[page_index]
    side_label = "Left" if side == "left" else "Right"

    data = [[
        Paragraph(f"{side_label} Page {page_index + 1} (No Match)", styles['Normal']),
    ]]

    if config.include_thumbnails and page.thumbnail:
        img = _pil_to_reportlab_image(page.thumbnail, 100*mm)
        data.append([img])

    table = Table(data, colWidths=[120*mm])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#f39c12')),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#dddddd')),
    ]))

    return table


def _pil_to_reportlab_image(pil_image: Image.Image, max_width: float):
    """Convert PIL Image to ReportLab Image."""
    from reportlab.platypus import Image as RLImage
    from reportlab.lib.units import mm

    if pil_image is None:
        return ""

    # Convert to RGB if necessary
    if pil_image.mode != 'RGB':
        pil_image = pil_image.convert('RGB')

    # Save to buffer
    buffer = io.BytesIO()
    pil_image.save(buffer, format='PNG')
    buffer.seek(0)

    # Calculate dimensions
    aspect = pil_image.height / pil_image.width
    width = min(max_width, 120*mm)
    height = width * aspect

    # Limit height
    max_height = 80*mm
    if height > max_height:
        height = max_height
        width = height / aspect

    return RLImage(buffer, width=width, height=height)
