import openpyxl as op
import re
import sys
import pandas as pd
from itertools import product
import datetime
import logging
import pathlib as pl
from parse import parse
import json

logger = logging.getLogger(__name__)
logging.basicConfig(filename='log.log', filemode='w', level=logging.INFO)


def detect_line_item(cell):
    LINE_ITEM_REGEX = "[0-9]*-[0-9]{2}"
    if not cell.value:
        return False
    else:
        return re.match(LINE_ITEM_REGEX, str(cell.value))


def extract_quarter(cell):
    assert type(cell.value) is str
    logging.info(cell.value)
    date = parse("{}{:ta}", cell.value)[1]
    logging.info(str(date))
    return date


def extract_revenues_and_expenses(filename, columns, info_row, info_column, name_cell, quarter_cell, sheet_names,
                                  first_column,
                                  line_items=pd.read_csv('Line_Items.csv', index_col=0)):
    print(filename)
    logging.info(filename)
    wb = op.load_workbook(filename, data_only=True)

    data = []
    for sheet_name in sheet_names:
        logging.info(f'{filename=}\t{sheet_name=}')
        try:
            sheet = wb[sheet_name]
        except KeyError as e:
            print(e)
            logging.warning(f'{filename}\t{sheet_name} didn\'t exist')
            continue
        name = sheet[name_cell].value
        quarter = extract_quarter(sheet[quarter_cell])
        data_rows = [cell.row for cell in sheet[info_column[0]]
                     if detect_line_item(cell)
                     ]
        data_columns = [cell.column for cell in sheet[info_row]
                        if cell.value
                        and cell.column >= first_column
                        and 'total' not in str(cell.value).lower()
                        ]
        for j in data_columns:
            for i in data_rows:
                if (cell := sheet.cell(i, j)).value or cell.value == 0:
                    data.append(
                        {
                            'Name': name,
                            'Value': cell.value,
                            'Quarter': quarter,
                            'Sheet': sheet_name,
                            'Column': sheet.cell(info_row, j).value,
                            'Line Item': (li := sheet.cell(i, info_column[1]).value),
                            'Line Category': line_items.loc[li]['Line Category'],
                            'Line Name': line_items.loc[li]['Line Name'],
                            'Revenue Expense Indicator': line_items.loc[li]['Revenue Expense Indicator']
                        })
    df = pd.DataFrame.from_records(data, columns=columns)
    return df, name


def main():
    filenames = sys.argv[1:]
    dfs = [extract_revenues_and_expenses(filename, **params) for filename in filenames]
    now = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    output_directory = pl.Path(f'Output/{program}/{now}')
    output_directory.mkdir(parents=True)
    logging.info(f'{output_directory=}')
    total_df = pd.DataFrame(columns=params['columns'])
    names = set([x[1] for x in dfs])
    for name in names:
        matching_dfs = [x[0] for x in dfs if x[1] == name]
        result = pd.DataFrame(columns=params['columns'])
        for matching_df in matching_dfs:
            result = result.append(matching_df)
            result.to_excel(output_directory / f'{name}.xlsx', index=False)
            total_df = total_df.append(result)
    total_df.to_excel(output_directory / 'Total.xlsx',
                      index=False)


if __name__ == '__main__':
    options = ['ACC', 'RBHA', 'ALTCS']
    program = None
    while program not in options:
        program = input(f"Please choose a program. Your options are: {', '.join(options)}\n\t")
    assert program in options
    with pl.Path(f'Input/formats/{program}.json').open('r') as f:
        params = json.load(f)
    try:
        main()
    except Exception as e:
        logging.exception(e)
        print(e)
        input()
        raise e
