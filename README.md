# python_excel2json

A simple way to parse an excel file and get the json that you want.

## Install

```
pip install python_excel2json
```

## Dependencies

```
xlrd
```

## Usage

```
from python_excel2json import parse_excel_to_json

# Step 1: Define the format of your excel that you want to parse.

excel_sheets_format = {
    'start_row_sheet_parsing': 1,
    'start_column_sheet_parsing': 0,
    'sheet_formats': [
        {
            'sheet_index': 0,
            'column_names': [
                {
                    'name': 'name1',
                    'type': 'str'
                },
                {
                    'name': 'name2',
                    'type': 'str'
                },
                {
                    'name': 'name3',
                    'type': 'str'
                },
                {
                    'name': 'name4',
                    'type': 'float'
                }
            ],
            'is_ordered': True
        },
        {
            'sheet_index': 1,
            'column_names': [
                {
                    'name': 'name1',
                    'type': 'str'
                },
                {
                    'name': 'name2',
                    'type': 'str'
                },
                {
                    'name': 'name3',
                    'type': 'str'
                },
                {
                    'name': 'name4',
                    'type': 'float'
                }
            ],
            'is_ordered': True,
            'start_row_sheet_parsing': 3,
            'start_column_sheet_parsing': 5
        }
    ]
}

# Step 2: Define your excel file input. Using 'xlrd' it can be from two ways:

file = '/my/path/my-file.xlsx'
result = parse_excel_to_json(excel_sheets_format, filename=file)

file = your-excel-content
result = parse_excel_to_json(excel_sheets_format, file_contents=file)

# Step 3: Obtain your specific json result.

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

```

## License

[MIT](./LICENSE)
