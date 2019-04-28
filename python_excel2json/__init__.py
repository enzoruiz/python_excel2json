from .main import (
    get_book_from_input_file, find_format_of_columns_from_specific_index_sheet
)


def parse_excel_to_json(
    excel_sheets_format, filename=None, file_contents=None
):
    """ It receives a file from a request and excel_sheets_format with
    the next format:

    excel_sheets_format = {
        'start_row_sheet_parsing': 1,
        'start_column_sheet_parsing': 0,
        'sheet_formats': [
            {
                'sheet_index': 0,
                'column_names': ['name1', 'name2', 'name3'],
                'is_ordered': True
            },
            {
                'sheet_index': 1,
                'column_names': ['name1', 'name2', 'name3', 'name4'],
                'is_ordered': True
            }
        ]
    }

    Then iterate inside each sheet and columns that you
    specify in the format.
    The response will be a list of list of sheets with objects inside
    like the next format.

    [
        {
            'sheet_index': 0,
            'results': [
                {
                    'name1': 'value',
                    'name2': 'value',
                    'name3': 'value'
                },
                ...
            ]
        },
        {
            'sheet_index': 1,
            'results': [
                {
                    'name1': 'value',
                    'name2': 'value',
                    'name3': 'value',
                    'name4': 'value'
                },
                ...
            ]
        }
    ]
    """

    # TODO: verify order of columns ('is_ordered')
    try:
        book = get_book_from_input_file(filename, file_contents)
        excel_parsed_list = []
        for sheet_format in excel_sheets_format.get('sheet_formats'):
            sheet_index = sheet_format.get('sheet_index')
            parsed_object = {}
            sheet = book.sheet_by_index(sheet_index)
            (
                sheet_format,
                start_row_sheet_parsing,
                start_column_sheet_parsing
            ) = find_format_of_columns_from_specific_index_sheet(
                excel_sheets_format, sheet_index
            )
            total_number_of_columns = len(sheet_format.get('column_names'))
            sheet_parsed_list = []

            for row_object in range(
                start_row_sheet_parsing, sheet.nrows
            ):
                row_parsed_dict = {}
                column_position_in_column_names = 0
                for column_position_in_sheet in range(
                    start_column_sheet_parsing, total_number_of_columns
                ):
                    row_parsed_dict[sheet_format.get('column_names')[
                        column_position_in_column_names
                    ]] = sheet.row(row_object)[column_position_in_sheet].value

                    column_position_in_column_names += 1
                sheet_parsed_list.append(row_parsed_dict)
            parsed_object['sheet_index'] = sheet_index
            parsed_object['results'] = sheet_parsed_list
            excel_parsed_list.append(parsed_object)
    except IndexError:
        raise Exception('Not enough data to parse the sheets from excel.')

    return excel_parsed_list