import sys, os, logging
import toml, openpyxl
from datetime import datetime
import ss_api

start_time = datetime.now()

CONFIG = None
_dir_in = os.path.join(os.path.dirname(os.path.abspath(__file__)), "in/")
_conf = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.toml")

def get_sheet(_dir_out):
    if isinstance(CONFIG, dict):
        print("Starting ...")
        if "verbose" in CONFIG and CONFIG["verbose"] == True:
            logging.basicConfig(filename="sheet.log", level=logging.INFO)

        for k, v in CONFIG["tables"].items():
            table_id = v["id"]  # This is the sheet_id
            table_src = v["src"]
            table_name = k

            # Call get_sheet_as_excel with sheet_id directly
            print("uploader", _dir_out)
            ss_api.get_sheet_as_excel(
                table_id,  # Pass the sheet_id
                os.path.join(_dir_out, f'{table_name}.xlsx')
            )

if __name__ == "__main__":
    with open(_conf, "r") as conf:
        CONFIG = toml.load(conf)
        if isinstance(CONFIG, dict):
            for k, v in CONFIG["env"].items():
                os.environ[k] = v
            print(CONFIG.get("verbose"))
            if CONFIG.get("verbose", False):
                logging.basicConfig(
                    filename="sheet.log", filemode="w", level=logging.INFO
                )

            # Check the command-line argument and call the appropriate function
            if sys.argv[1] == "get":
                # Ensure the directory is passed as an argument
                if len(sys.argv) > 2:
                    dir_out = sys.argv[2]
                    print("47", dir_out)
                    get_sheet(dir_out)
                else:
                    print("Error: Output directory not specified.")
            elif sys.argv[1] == "set":
                set_sheet()
            elif sys.argv[1] == "attach":
                attach_sheet()
            elif sys.argv[1] == "update":
                update_sheet()
            elif sys.argv[1] == "summary":
                make_summary()

    if isinstance(CONFIG, dict):
        with open(_conf, "w") as conf:
            toml.dump(CONFIG, conf)

end_time = datetime.now()
print("Duration: {}".format(end_time - start_time))
