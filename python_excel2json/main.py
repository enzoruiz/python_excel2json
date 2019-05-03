from xlrd import open_workbook


DEFAULT_START_ROW_EXCEL_PARSING = 1
DEFAULT_START_COLUMN_EXCEL_PARSING = 0

authorized_types = {
    'int': int, 'float': float, 'str': str
}


def get_book_from_input_file(filename=None, file_contents=None):
    if not filename and not file_contents:
        raise Exception('Please insert an excel file.')

    if filename:
        book = open_workbook(filename=filename)

    if file_contents:
        book = open_workbook(file_contents=file_contents)

    return book


def find_format_of_columns_from_specific_index_sheet(
    excel_sheets_format, sheet_index
):
    """ Find the data format, start row index and start column index from
    the excel_sheets_format using the actual sheet index.
    """
    column_format_found = {}
    start_row_sheet_parsing = excel_sheets_format.get(
        'start_row_sheet_parsing'
    )
    start_column_sheet_parsing = excel_sheets_format.get(
        'start_column_sheet_parsing'
    )
    for sheet_format in excel_sheets_format.get('sheet_formats'):
        if sheet_format.get('sheet_index') == sheet_index:
            column_format_found = sheet_format
            format_start_row_sheet_parsing = sheet_format.get(
                'start_row_sheet_parsing'
            )
            if format_start_row_sheet_parsing:
                start_row_sheet_parsing = format_start_row_sheet_parsing
            format_start_column_sheet_parsing = sheet_format.get(
                'start_column_sheet_parsing'
            )
            if format_start_column_sheet_parsing:
                start_column_sheet_parsing = format_start_column_sheet_parsing
            break

    start_row_sheet_parsing = (
        DEFAULT_START_ROW_EXCEL_PARSING if not start_row_sheet_parsing
        else start_row_sheet_parsing
    )
    start_column_sheet_parsing = (
        DEFAULT_START_COLUMN_EXCEL_PARSING if not start_column_sheet_parsing
        else start_column_sheet_parsing
    )

    return (
        column_format_found, start_row_sheet_parsing,
        start_column_sheet_parsing
    )


def get_expected_value(value, expected_type):
    if type(value) == float and expected_type == 'str':
        value = str(value).replace('.0', '')

    value = authorized_types[expected_type](value) if value else value

    return value
