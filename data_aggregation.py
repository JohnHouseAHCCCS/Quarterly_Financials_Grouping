import pandas as pd
import sys
import pathlib as pl
import logging
import datetime

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(logging.StreamHandler())


def main():
    output_directory = pl.Path(f'Output/Results')
    now = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    output_file = output_directory / f'{now} {input("Filename? ")}.csv'
    files = sys.argv[1:]
    files = [pl.Path(file) for file in files]
    dfs = []  # type: list[pd.DataFrame]
    for file in files:
        dfs.append(pd.read_csv(file))
    total_df = dfs[0]
    for df in dfs[1:]:
        total_df = total_df.append(df)
    columns = total_df.columns
    try:  # This is to handle CHP files
        total_df = total_df.sort_values('Source')
        total_df.drop('Source', axis=1)
    except KeyError:
        pass
    except ValueError:
        pass
    total_df = total_df.drop_duplicates(columns)
    total_df.to_csv(output_file, index=False)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logger.exception(e)
        input()
        raise e
