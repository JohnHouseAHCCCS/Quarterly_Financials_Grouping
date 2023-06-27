import pathlib as pl
import logging
import json
import sys
import pyperclip as pc

# region Logging
LOGFOLDER = pl.Path('Logs')
FILENAME = f'{pl.Path(__file__).stem}.log'
# FORMAT = "%(levelname) - %(funcName)s - %(message)s"
FILE_LEVEL = logging.DEBUG
STREAM_LEVEL = logging.DEBUG
LOG_LEVEL = logging.DEBUG
logger = logging.getLogger(__name__)
stream_handler = logging.StreamHandler()
stream_handler.setLevel(STREAM_LEVEL)
logger.addHandler(stream_handler)
if not LOGFOLDER.exists():
    LOGFOLDER.mkdir()
logfile = LOGFOLDER / FILENAME
file_handler = logging.FileHandler(logfile, mode='w')
# formatter = logging.Formatter(FORMAT)
# file_handler.setFormatter(formatter)
file_handler.setLevel(FILE_LEVEL)
logger.addHandler(file_handler)
logger.setLevel(LOG_LEVEL)
# endregion

# region Parameters
with open('programs.json', 'r') as f:
    programs = json.load(f)
    programs = [key for key, val in programs.items() if val]
    logger.debug(programs)
output_directory = pl.Path('Input')
# endregion


def move_program(program_directory: pl.Path, program):
    target = output_directory / program
    for file in program_directory.iterdir():
        name = file.name
        if 'ytd' in name.lower():
            logger.debug('Skipped {name}')
            continue
        elif (target / name).exists():
            continue
        else:
            logger.info(f'Moved {name}')
            file.rename(target / name)


def move_quarter(quarter_directory: pl.Path):
    for program_directory in quarter_directory.iterdir():
        if program_directory.is_dir():
            logger.info(program_directory)
            for program in programs:
                if program in program_directory.name:
                    move_program(program_directory, program)
                    logger.debug(f'Moved {program_directory}')


def main():
    directory = pl.Path(pc.paste())
    for quarter_directory in directory.iterdir():
        if quarter_directory.is_dir():
            logger.debug(quarter_directory)
            move_quarter(quarter_directory)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logger.exception(e)
        raise e
