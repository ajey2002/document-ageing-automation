from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def setup_table_styles():
    header_fill = PatternFill("solid", fgColor="BDD7EE")
    header_font = Font(color="000000", bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    title_font = Font(bold=True, size=14)
    title_align = Alignment(horizontal="center", vertical="center")

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    total_font = Font(bold=True)

    return {
        "header_fill": header_fill,
        "header_font": header_font,
        "header_align": header_align,
        "title_font": title_font,
        "title_align": title_align,
        "border": border,
        "total_font": total_font,
    }
