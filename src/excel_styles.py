from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def setup_table_styles():
    # Header
    header_fill = PatternFill("solid", fgColor="9DC3E6") 
    header_font = Font(color="000000", bold=True)
    header_align = Alignment(horizontal="center", vertical="bottom", wrap_text=True)

    title_font = Font(bold=True, size=14)
    title_align = Alignment(horizontal="center", vertical="bottom")

    # Default alignments
    bottom_left = Alignment(horizontal="left", vertical="bottom", wrap_text=False)
    bottom_center = Alignment(horizontal="center", vertical="bottom", wrap_text=False)
    bottom_right = Alignment(horizontal="right", vertical="bottom", wrap_text=False)

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Double top border for totals (SAP-ish)
    double = Side(style="double", color="000000")
    total_top_border = Border(left=thin, right=thin, top=double, bottom=thin)

    total_font = Font(bold=True)

    return {
        "header_fill": header_fill,
        "header_font": header_font,
        "header_align": header_align,
        "title_font": title_font,
        "title_align": title_align,
        "border": border,
        "total_top_border": total_top_border,
        "total_font": total_font,
        "bottom_left": bottom_left,
        "bottom_center": bottom_center,
        "bottom_right": bottom_right,
    }
