import logging
import argparse
import find_origin
from find_origin import process_flow
import pandas as pd
import configparser

def setup_logging():
    parser = argparse.ArgumentParser(description="Program description")
    log_group = parser.add_mutually_exclusive_group()
    log_group.add_argument("--debug", action="store_const", dest="loglevel", const=logging.DEBUG,
                           help="Set logging to DEBUG level")
    log_group.add_argument("--info", action="store_const", dest="loglevel", const=logging.INFO,
                           help="Set logging to INFO level")
    log_group.add_argument("--logging", action="store_const", dest="loglevel", const=logging.WARNING,
                           help="Set logging to WARNING level")
    log_group.add_argument("--quiet", action="store_const", dest="loglevel", const=logging.ERROR,
                           help="Silence warnings and show only errors")
    parser.set_defaults(loglevel=logging.WARNING)
    parser.add_argument("--test", action="store_true", help="Run unit tests")
    parser.add_argument("--lot", type=int, help="Specify the lot number")
    parser.add_argument("--config", type=str, help="Path to configuration file (configconservatoire.ini)")
    args = parser.parse_args()
    logging.getLogger().setLevel(args.loglevel)

    return args

def read_config(config_path):
    config = configparser.ConfigParser()
    config.read(config_path)
    return config

# In your main function or script
def main():
    args = setup_logging()
    if not args.config:
        raise Exception('a config file is required')
    config = read_config(args.config)
    find_origin.GRANDE_COLLECTION_PATH = config['Paths']['GRANDE_COLLECTION_PATH']
    find_origin.SEMIS_ANNEES_ANTERIEURES_PATH = config['Paths']['SEMIS_ANNEES_ANTERIEURES_PATH']
    find_origin.JARDIN_PLANTES_MENACEES_PATH = config['Paths']['JARDIN_PLANTES_MENACEES_PATH']
    find_origin.GC_FILENAME = config['Files']['GC_FILENAME']
    find_origin.PATH_SEP = config['Paths']['PATH_SEP']
    flow_fname = config['Files']['FLOW_FILENAME']
    sheet_name = config['Files']['DEFAULT_SHEET_NAME']
    if args.test:
        sheet_name='Unit Tests'
    df = pd.read_excel(flow_fname, sheet_name=sheet_name, header=1)

    # logging.info("This is an info message")
    # logging.warning("This is a warning message")
    # logging.debug("This is a debug message")

    process_flow(df, args.lot, args.test)
    output_sheet = 'updated'
    try:
        with pd.ExcelWriter(flow_fname, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name=output_sheet, index=False)
    except ValueError:
        logging.debug("writing to sheet failed. Make sure to remove the updated sheet before running")
        output_sheet += '_'
        with pd.ExcelWriter(flow_fname, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name=output_sheet, index=False)

if __name__ == "__main__":
    main()
