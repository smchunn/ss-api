import os, logging
from typing import Dict
import httpx, truststore, ssl
from pathlib import Path
from urllib.parse import urlencode
import time


class APIException(Exception):
    def __init__(self, message, response):
        super().__init__(message)
        self.response = response


def retry_on_500(max_retries=3, delay=2):
    def decorator(func):
        def wrapper(*args, **kwargs):
            retries = 0
            while True:
                try:
                    return func(*args, **kwargs)
                except APIException as e:
                    if e.response.status_code == 500:
                        retries += 1
                        if retries > max_retries:
                            logging.error(f"Max retries exceeded for {func.__name__}")
                            raise
                        logging.warning(f"500 error on {func.__name__}, retrying {retries}/{max_retries} after {delay}s...")
                        time.sleep(delay)
                    else:
                        raise
        return wrapper
    return decorator


def list_sheets(*, access_token=None) -> None | Dict:
    """
    Retrieve all Smartsheet sheets accessible to the user.

    Args:
        access_token (str, optional): Bearer token. If not provided, uses the SMARTSHEET_ACCESS_TOKEN environment variable.

    Returns:
        dict: Sheets list on success, or None on failure.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            params = urlencode({"includeAll": "true"})
            url = f"https://api.smartsheet.com/2.0/sheets?{params}"
            headers = {
                "Authorization": f"Bearer {bearer}",
            }
            logging.info(f"GET: list sheets, {url},{headers}")
            response = client.get(
                url=url,
                headers=headers,
                timeout=60,
            )
            if response.status_code != 200:
                raise APIException(f"GET: list sheets, {url},{headers}", response)
            return response.json()
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")

    return None


@retry_on_500(max_retries=3, delay=2)
def create_sheet(folder_id, sheet, *, access_token=None) -> None | Dict:
    """
    Create a new sheet inside a specified folder.

    Args:
        folder_id (str): Folder ID to create the sheet in.
        sheet (dict): Sheet definition as per Smartsheet API.
        access_token (str, optional): Bearer token.

    Returns:
        dict: Response dict on success, or None on failure.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/folders/{folder_id}/sheets"
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Type": "application/json",
            }
            logging.info(f"POST: create sheet, {url},{headers}, payload: {sheet}")
            response = client.post(
                url=url,
                headers=headers,
                json=sheet,
                timeout=60,
            )
            if response.status_code != 200:
                raise APIException(f"POST: create sheet, {url},{headers}", response)
            return response.json()
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")

    return None


def get_sheet(sheet_id, last_modified=None, *, access_token=None) -> None | Dict:
    """
    Retrieve details of a sheet by ID.

    Args:
        sheet_id (str): Sheet ID.
        last_modified (str, optional): ISO date string; only rows modified since will be included.
        access_token (str, optional): Bearer token.

    Returns:
        dict: Sheet data on success, None if not found or on error.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}"
            if last_modified:
                params = urlencode(
                    {"rowsModifiedSince": last_modified, "include": "writerInfo"}
                )
                url += f"?{params}"
            headers = {
                "Authorization": f"Bearer {bearer}",
            }
            logging.info(f"GET: get sheet, {url},{headers}")
            response = client.get(
                url=url,
                headers=headers,
                timeout=60,
            )
            if response.status_code == 404:
                # Sheet does not exist, return None
                return None
            if response.status_code != 200:
                # Handle other errors
                raise APIException(f"GET: get sheet, {url},{headers}", response)
            return response.json()
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")
    return None


def get_sheet_as_xlsx(sheet_id, filepath, *, access_token=None):
    """
    Download a sheet as an Excel file (XLSX).

    Args:
        sheet_id (str): Sheet ID.
        filepath (str): Local file path to save the .xlsx.
        access_token (str, optional): Bearer token.

    Returns:
        dict: Sheet data on success, None on error.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}"
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Accept": "application/vnd.ms-excel",
            }
            response = client.get(
                url=url,
                headers=headers,
                timeout=60,
            )
            if response.status_code != 200:
                try:
                    error_msg = response.json()
                except Exception:
                    error_msg = response.text
                raise APIException(
                    f"GET: get sheet, {url},{headers}\nError: {error_msg}", response
                )

            with open(filepath, "wb") as f:
                f.write(response.content)
            print(f"File saved as {filepath}")
            return filepath
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")

    return None


@retry_on_500(max_retries=3, delay=2)
def update_sheet(sheet_id, updates, *, access_token=None, batch_size=100):
    """
    Batch update rows in a sheet.

    Args:
        sheet_id (str): Sheet ID.
        updates (list of dict): Row update dictionaries.
        access_token (str, optional): Bearer token.
        batch_size (int, optional): Number of rows per batch.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        sheet = get_sheet(sheet_id, access_token=bearer)
        if not sheet or not sheet["rows"]:
            return

        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows"
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Type": "application/json",
            }

            for i in range(0, len(updates), batch_size):
                json = updates[i : i + batch_size]
                response = client.put(
                    url=url,
                    headers=headers,
                    json=json,
                    timeout=60,
                )
                logging.info(
                    f"PUT: update sheet, url:{url},headers:{headers},json:{json}"
                )
                if response.status_code != 200:
                    raise APIException(f"PUT: update rows, {url},{headers}", response)
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")


@retry_on_500(max_retries=3, delay=2)
def move_rows(target_sheet_id, source_sheet_id, *, access_token=None):
    """
    Move all rows from one sheet to another.

    Args:
        target_sheet_id (str): Destination sheet ID.
        source_sheet_id (str): Source sheet ID.
        access_token (str, optional): Bearer token.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        source_sheet = get_sheet(source_sheet_id, access_token=bearer)
        rows = []
        if not source_sheet:
            return
        for row in source_sheet["rows"]:
            rows.append(row["id"])

        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/sheets/{source_sheet_id}/rows/move"
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Type": "application/json",
            }

            batch = 200
            for i in range(0, len(rows), batch):
                response = client.post(
                    url=url,
                    headers=headers,
                    json={
                        "rowIds": rows[i : i + batch],
                        "to": {"sheetId": target_sheet_id},
                    },
                    timeout=120,
                )
                if response.status_code != 200:
                    raise APIException(
                        f"POST: move all rows, {url},{headers}", response
                    )
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")


@retry_on_500(max_retries=3, delay=2)
def add_rows(sheet_id, rows, *, access_token=None, batch_size=500):
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        sheet = get_sheet(sheet_id, access_token=bearer)
        if not sheet:
            return

        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows"
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Type": "application/json",
            }

            for i in range(0, len(rows), batch_size):
                json = rows[i : i + batch_size]
                response = client.post(
                    url=url,
                    headers=headers,
                    json=json,
                    timeout=60,
                )
                logging.info(f"POST: add_rows, url:{url},headers:{headers},json:{json}")
                if response.status_code != 200:
                    raise APIException(f"POST: add_rows, {url},{headers}", response)
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")


@retry_on_500(max_retries=3, delay=2)
def delete_rows(sheet_id, rows, *, access_token=None):
    print("reached")
    try:
        responses = []
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            batch_size = 100
            for i in range(0, len(rows), batch_size):
                batch = ",".join(str(row) for row in rows[i : i + batch_size])
                url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows?ids={batch}&ignoreRowsNotFound=true"
                headers = {
                    "Authorization": f"Bearer {bearer}",
                }
                response = client.delete(
                    url=url,
                    headers=headers,
                    timeout=60,
                )
                responses.append(response)
                if response.status_code != 200:
                    raise APIException(
                        f"DELETE: delete rows, {url},{headers}", response
                    )
            return responses
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")
    return None


@retry_on_500(max_retries=3, delay=2)
def delete_all_rows(sheet_id, *, access_token=None):
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        # Retrieve the sheet to ensure it exists and to get the rows
        sheet = get_sheet(sheet_id, access_token=bearer)
        if not sheet:
            print("Sheet not found or could not be retrieved.")
            return

        # Check if there are rows to delete
        if not sheet.get("rows"):
            print("No rows to delete.")
            return

        # Collect all row IDs for deletion
        row_ids = [row["id"] for row in sheet["rows"]]
        batch_size = 100  # Smartsheet API allows deleting up to 100 rows at a time

        # Delete rows in batches
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            for i in range(0, len(row_ids), batch_size):
                batch = ",".join(str(row_id) for row_id in row_ids[i : i + batch_size])
                delete_url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows?ids={batch}&ignoreRowsNotFound=true"
                headers = {
                    "Authorization": f"Bearer {bearer}",
                }
                response = client.delete(
                    url=delete_url,
                    headers=headers,
                    timeout=60,
                )
                if response.status_code != 200:
                    raise APIException(
                        f"DELETE: delete rows, {delete_url}, {headers}", response
                    )

    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")

    return None


@retry_on_500(max_retries=3, delay=2)
def delete_sheet(sheet_id, *, access_token=None):
    bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
    ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    with httpx.Client(verify=ssl_context) as client:
        try:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}"
            headers = {
                "Authorization": f"Bearer {bearer}",
            }
            response = client.delete(
                url=url,
                headers=headers,
                timeout=60,
            )
            if response.status_code != 200:
                raise APIException(f"GET: get sheet, {url},{headers}", response)
            return response.json()
        except APIException as e:
            logging.error(f"API Error: {e.response}")
            print(f"An error occurred: {e.response}")

    return None


@retry_on_500(max_retries=3, delay=2)
def clear_sheet(sheet_id, *, access_token=None):
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        sheet = get_sheet(sheet_id, access_token=bearer)
        if not sheet:
            exit()

        if not sheet["rows"]:
            return

        first_row_id = sheet["rows"][0]["id"]
        data = [
            {"id": row["id"], "parentId": first_row_id}
            for i, row in enumerate(sheet["rows"])
            if i > 0
        ]
        update_sheet(sheet_id, data, access_token=bearer)
        delete_rows(sheet_id, [first_row_id], access_token=bearer)
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")


@retry_on_500(max_retries=3, delay=2)
def import_xlsx_sheet(
    sheet_name,
    filepath,
    folder_id=None,
    primary_column_index=0,
    *,
    access_token=None,
    timeout=240,
):
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client, open(filepath, "br") as xl:
            if folder_id:
                params = urlencode(
                    {
                        "sheetName": sheet_name,
                        "headerRowIndex": 0,
                        "primaryColumnIndex": primary_column_index,
                    }
                )
                url = f"https://api.smartsheet.com/2.0/folders/{folder_id}/sheets/import?{params}"
            else:
                params = urlencode(
                    {
                        "sheetName": sheet_name,
                        "headerRowIndex": 0,
                        "primaryColumnIndex": primary_column_index,
                    }
                )
                url = f"https://api.smartsheet.com/2.0/sheets/import?{params}"

            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Disposition": "attachment",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
            response = client.post(
                url=url,
                headers=headers,
                content=xl,
                timeout=timeout,
            )
            if response.status_code != 200:
                raise APIException("POST: import excel", response)
            return response.json()
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")
    return None


def attach_file(sheet_id, filepath, *, access_token=None):
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client, open(filepath, "br") as xl:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/attachments"
            file_size = Path(filepath).stat().st_size
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Type": "application/vnd.ms-excel",
                "Content-Disposition": f'attachment; filename="{os.path.basename(filepath)}"',
                "Content-Length": str(file_size),
            }
            response = client.post(
                url=url,
                headers=headers,
                content=xl,
                timeout=60,
            )
            if response.status_code != 200:
                raise APIException(f"POST: attach file, {url},{headers}", response)
            return response
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")
    return None


def get_columns(sheet_id, *, access_token=None):
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/columns"
            headers = {
                "Authorization": f"Bearer {bearer}",
            }
            logging.info(f"GET: get columns, {url},{headers}")
            response = client.get(
                url=url,
                headers=headers,
                timeout=60,
            )
            if response.status_code == 404:
                # Sheet does not exist, return None
                return None
            if response.status_code != 200:
                # Handle other errors
                raise APIException(f"GET: get columns, {url},{headers}", response)
            return response.json()
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")

    return None


def update_columns(sheet_id, column_id, column_update, *, access_token=None):
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = (
                f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/columns/{column_id}"
            )
            headers = {
                "Authorization": f"Bearer {bearer}",
            }
            logging.info(f"PUT: update columns, {url}, payload: {column_update}")
            response = client.put(
                url=url,
                headers=headers,
                json=column_update,
                timeout=60,
            )
            if response.status_code == 404:
                # Sheet does not exist, return None
                return None
            if response.status_code != 200:
                # Log full response content for debugging
                try:
                    error_content = response.json()
                except Exception:
                    error_content = response.text
                logging.error(f"PUT: update columns failed, response: {error_content}")
                raise APIException(
                    f"PUT: update columns, {url},{headers},{column_update}", response
                )
            return response.json()
    except APIException as e:
        logging.error(f"API Error: {e.response.json()}")
        print(f"An error occurred: {e.response}")

    return None


def rename_sheet(sheet_id, new_sheet_name, *, access_token=None):
    bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
    ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    with httpx.Client(verify=ssl_context) as client:
        try:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}"
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Type": "application/json",  # Set the content type to JSON
            }
            # Create the payload with the new sheet name
            payload = {"name": new_sheet_name}
            response = client.put(
                url=url,
                headers=headers,
                json=payload,  # Use json to send the payload
                timeout=60,
            )
            if response.status_code != 200:
                raise APIException(
                    f"PUT: rename sheet, {url}, {headers}, {payload}", response
                )
            return response.json()
        except APIException as e:
            logging.error(f"API Error: {e.response}")
            print(f"An error occurred: {e.response}")

    return None


def update_column_description(sheet_id, column_id, new_description, *, access_token=None):
    """
    Update the description of a column in a sheet.

    Args:
        sheet_id (str): The ID of the sheet.
        column_id (str): The ID of the column to update.
        new_description (str): The new description text for the column.
        access_token (str, optional): Bearer token.

    Returns:
        dict: Updated column data on success, None on failure.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        logging.info(f"PUT: update column description, sheet_id={sheet_id}, column_id={column_id}")
        column_update = {
            "description": new_description
        }
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/columns/{column_id}"
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Type": "application/json",
            }
            response = client.put(
                url=url,
                headers=headers,
                json=column_update,
                timeout=60,
            )
            if response.status_code == 404:
                logging.error(f"Column or sheet not found: sheet_id={sheet_id}, column_id={column_id}")
                return None
            if response.status_code != 200:
                logging.error(f"Failed to update column description: {response.text}")
                print(f"An error occurred: {response.text}")
                return None
            logging.info(f"Successfully updated column description: sheet_id={sheet_id}, column_id={column_id}")
            return response.json()
    except Exception as e:
        logging.error(f"Exception during update_column_description: {e}")
        print(f"An error occurred while updating column description: {e}")
        return None


@retry_on_500(max_retries=3, delay=2)
def add_sheet_collaborator(sheet_id, email, access_level="EDITOR", *, access_token=None):
    """
    Add a collaborator to a sheet.

    Args:
        sheet_id (str): The ID of the sheet.
        email (str): The email address of the collaborator to add.
        access_level (str, optional): Access level for the collaborator. Defaults to "EDITOR".
        access_token (str, optional): Bearer token.

    Returns:
        dict: Response data on success, None on failure.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        with httpx.Client(verify=ssl_context) as client:
            url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/shares"
            headers = {
                "Authorization": f"Bearer {bearer}",
                "Content-Type": "application/json",
            }
            payload = {
                "email": email,
                "accessLevel": access_level
            }
            response = client.post(
                url=url,
                headers=headers,
                json=payload,
                timeout=60,
            )
            if response.status_code != 200:
                logging.error(f"Failed to add collaborator: {response.text}")
                print(f"An error occurred: {response.text}")
                return None
            return response.json()
    except Exception as e:
        logging.error(f"Exception during add_sheet_collaborator: {e}")
        print(f"An error occurred while adding collaborator: {e}")
        return None


def create_cross_sheet_reference(destination_sheet_id: int, source_sheet_id: int, source_column_id: int, name: str, *, access_token=None) -> dict | None:
    """
    Create a cross-sheet reference in the destination sheet referencing a range from the source sheet.

    Returns:
        dict: API response on success, None on failure.
    """
    try:
        bearer = access_token or os.environ["SMARTSHEET_ACCESS_TOKEN"]
        ssl_context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        url = f"https://api.smartsheet.com/2.0/sheets/{destination_sheet_id}/crosssheetreferences"
        headers = {
            "Authorization": f"Bearer {bearer}",
            "Content-Type": "application/json",
        }
        payload = {
            "name": name,
            "sourceSheetId": source_sheet_id,
            "startColumnId": source_column_id,
            "endColumnId": source_column_id,
        }
        with httpx.Client(verify=ssl_context) as client:
            response = client.post(url, headers=headers, json=payload, timeout=60)
            if response.status_code != 200:
                logging.error(f"Failed to create cross-sheet reference: {response.text}")
                print(f"An error occurred: {response.text}")
                return None
            return response.json()
    except APIException as e:
        logging.error(f"API Error: {e.response}")
        print(f"An error occurred: {e.response}")
        return None
