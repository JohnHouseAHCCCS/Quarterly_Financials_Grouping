import pandas as pd
import pathlib as pl
import logging
import json

# region Logging
LOGFOLDER = pl.Path('Logs')
FILENAME = f'{pl.Path(__file__).stem}.log'
FORMAT = "%(asctime)s - %(message)s"
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(logging.StreamHandler())
if not LOGFOLDER.exists():
    LOGFOLDER.mkdir()
logfile = LOGFOLDER / FILENAME
handler = logging.FileHandler(logfile, mode='w')
formatter = logging.Formatter(FORMAT)
handler.setFormatter(formatter)
logger.addHandler(handler)
# endregion


# region Parameters
with open('programs.json', 'r') as f:
    programs = json.load(f)
    programs = [key for key, val in programs.items() if val]
output_directory = pl.Path(f'Output/Results')
# endregion


def aggregate_dataframes(program):
    files = (pl.Path('Output') / program).iterdir()
    output_file = output_directory / f'{program}.csv'
    dfs = []  # type: list[pd.DataFrame]
    for file in files:
        logger.info(file)
        dfs.append(pd.read_csv(file))
    total_df = pd.concat(dfs)
    try:  # This is to handle CHP files
        total_df = total_df.sort_values('Source')
        total_df = total_df.drop('Source', axis=1)
    except KeyError:
        pass
    except ValueError:
        pass
    total_df = total_df.drop_duplicates()
    total_df.to_csv(output_file, index=False)


def main():
    for program in programs:
        aggregate_dataframes(program)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logger.exception(e)
        input()
        raise e
