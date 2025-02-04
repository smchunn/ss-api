import sys, os, logging
import toml, openpyxl
from datetime import datetime
from src import ss_api


start_time = datetime.now()

CONFIG = None
_dir_in = os.path.join(os.path.dirname(os.path.abspath(__file__)), "in/")
_dir_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out/")
_conf = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.toml")


def get_sheet():

    if isinstance(CONFIG, dict):
        print("Starting ...")
        if "verbose" in CONFIG and CONFIG["verbose"] == True:
            logging.basicConfig(filename="sheet.log", level=logging.INFO)
        for k, v in CONFIG["tables"].items():
            table_id = v["id"]
            table_src = v["src"]
            table_name = k
            ss_api.get_sheet_as_excel(table_id, _dir_out)

            workbook = openpyxl.load_workbook(os.path.join(_dir_out, table_src))
            worksheet = workbook[table_name]
            worksheet.title = "AUDIT"
            workbook.save(os.path.join(_dir_out, table_src))


def attach_sheet():

    if isinstance(CONFIG, dict):
        print("Starting ...")
        if "verbose" in CONFIG and CONFIG["verbose"] == True:
            logging.basicConfig(filename="sheet.log", level=logging.INFO)
        for k, v in CONFIG["tables"].items():
            table_id = v["id"]
            table_src = v["src"]
            table_name = k
            ss_api.attach_file(table_id, os.path.join(_dir_in, table_src))


def set_sheet():
    print("Starting ...")

    if isinstance(CONFIG, dict):
        target_folder_id = CONFIG["target_folder"]
        for k, v in CONFIG["tables"].items():
            table_id = v["id"]
            table_src = v["src"]
            table_name = k
            print(f"starting {table_name}...")

            if not table_id:
                print(f"No existing table, uploading {table_src} to {table_name}")
                result = ss_api.import_xlsx_sheet(
                    f"{table_name}",
                    os.path.join(_dir_in, table_src),
                    folder_id=target_folder_id,
                )
                if result:
                    table_id = str(result["result"]["id"])
                    print(f"  {table_name}({table_id}): new table loaded")
                    CONFIG["tables"][k]["id"] = table_id
            else:
                result = ss_api.import_xlsx_sheet(
                    f"TMP_{table_name}", os.path.join(_dir_in, table_src)
                )
                if not result:
                    continue

                if "message" in result and result["message"] != "SUCCESS":
                    print(result["message"])
                    return

                import_sheet_id = result["result"]["id"]
                target_sheet_id = table_id

                if not import_sheet_id or not target_sheet_id:
                    continue

                ss_api.clear_sheet(target_sheet_id)
                ss_api.move_rows(target_sheet_id, import_sheet_id)
                ss_api.delete_sheet(import_sheet_id)
            print("done...")


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

            if sys.argv[1] == "get":
                get_sheet()
            elif sys.argv[1] == "set":
                set_sheet()
            elif sys.argv[1] == "attach":
                set_sheet()
            elif sys.argv[1] == "test":
                attach_sheet()

    if isinstance(CONFIG, dict):
        with open(_conf, "w") as conf:
            toml.dump(CONFIG, conf)

end_time = datetime.now()
print("Duration: {}".format(end_time - start_time))
