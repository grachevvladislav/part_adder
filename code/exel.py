from .constants import MAIN_HEADER
from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
from .json_parser import Server
from openpyxl.worksheet.worksheet import Worksheet


def set_style(cell, header: bool = False, border_only: bool = False) -> None:
    cell.border = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000"),
    )
    if not border_only:
        if header:
            cell.alignment = Alignment(
                wrap_text=True,
                vertical="top",
            )
            cell.font = Font(
                name="Helvetica Neue (Headings)",
                size=13,
                bold=True,
                color="FF000000",
            )
        else:
            cell.alignment = Alignment(
                wrap_text=True,
                vertical="center",
                horizontal="center",
            )
            cell.font = Font(
                name="Helvetica Neue (Headings)", size=13, color="FF000000"
            )
            cell.fill = PatternFill(fill_type="solid", start_color="DCE2F1")


def fill_line():
    pass


def create_main_table(sheet: Worksheet, servers: list[Server]) -> None:
    # list zoom
    sheet.sheet_view.zoomScale = 55
    # header
    sheet.row_dimensions[1].height = 55
    for col, title in zip(
        sheet.iter_cols(max_col=len(MAIN_HEADER), max_row=1),
        MAIN_HEADER.items(),
    ):
        sheet.column_dimensions[col[0].column_letter].width = title[1]
        col[0].value = title[0]
        set_style(col[0], header=True)
    # table body
    line_cursor = 2
    for server in servers:
        for row, component in zip(
            sheet.iter_rows(
                max_col=len(MAIN_HEADER),
                min_row=line_cursor,
                max_row=line_cursor + len(server.config)
            ),
            server.config.values(),
        ):
            sheet.row_dimensions[row[0].row].height = 35
            # first line information

            values = [*['' for i in range(4)], *component.get_list(), '']
            for cell, field in zip(row, values):
                cell.value = field
                set_style(cell)
        # blank line
        sheet.merge_cells(
            start_row=line_cursor + len(server.config),
            start_column=1,
            end_row=line_cursor + len(server.config),
            end_column=len(MAIN_HEADER)
        )
        for col in sheet.iter_cols(
                max_col=len(MAIN_HEADER),
                min_row=line_cursor+len(server.config)
        ):
            set_style(col[0], border_only=True)
        sheet.row_dimensions[line_cursor + len(server.config)].height = 35
        line_cursor += len(server.config) + 1
