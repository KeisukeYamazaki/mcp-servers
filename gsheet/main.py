#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import os
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple, Union

import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from mcp.server.fastmcp import Context, FastMCP
from pydantic import BaseModel

# If modifying these SCOPES, delete your previously saved token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Create an MCP server
mcp = FastMCP("GoogleSheetsServer")


def get_credentials():
    """Get and refresh Google API credentials."""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_info(json.load(open("token.json")))

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Check if we have a service account key
            if os.path.exists("service_account.json"):
                creds = ServiceAccountCredentials.from_service_account_file("service_account.json", scopes=SCOPES)
            else:
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())

    return creds


def get_services():
    """Get Google Drive and Sheets services."""
    creds = get_credentials()
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    return drive_service, sheets_service


# ---- File operations ----


@mcp.tool()
def list_spreadsheets(folder_id: Optional[str] = None) -> str:
    """
    List all Google Sheets files, or those in a specific folder if folder_id is provided.

    Args:
        folder_id: Optional ID of the folder to list spreadsheets from

    Returns:
        JSON string with spreadsheet information
    """
    drive_service, _ = get_services()

    # Set up the query to find only Google Sheets
    query = "mimeType='application/vnd.google-apps.spreadsheet'"

    # Add folder constraint if provided
    if folder_id:
        query += f" and '{folder_id}' in parents"

    results = []
    page_token = None

    while True:
        try:
            response = (
                drive_service.files()
                .list(
                    q=query,
                    spaces="drive",
                    fields="nextPageToken, files(id, name, createdTime, modifiedTime, webViewLink)",
                    pageToken=page_token,
                )
                .execute()
            )

            files = response.get("files", [])
            results.extend(files)

            page_token = response.get("nextPageToken", None)
            if page_token is None:
                break

        except HttpError as error:
            return f"Error: {error}"

    return json.dumps(results, indent=2)


@mcp.tool()
def copy_spreadsheet(file_id: str, new_name: str, destination_folder_id: Optional[str] = None) -> str:
    """
    Copy a Google Sheets file with a new name, optionally to a destination folder.

    Args:
        file_id: ID of the spreadsheet to copy
        new_name: Name for the copy
        destination_folder_id: Optional ID of the destination folder

    Returns:
        JSON string with information about the new copy
    """
    drive_service, _ = get_services()

    # First create a copy
    body = {"name": new_name}

    try:
        copied_file = drive_service.files().copy(fileId=file_id, body=body).execute()

        # If destination folder specified, move the file there
        if destination_folder_id:
            # Get the current parents
            file = drive_service.files().get(fileId=copied_file["id"], fields="parents").execute()
            previous_parents = ",".join(file.get("parents"))

            # Move to the new folder
            drive_service.files().update(
                fileId=copied_file["id"],
                addParents=destination_folder_id,
                removeParents=previous_parents,
                fields="id, parents",
            ).execute()

        return json.dumps(copied_file, indent=2)

    except HttpError as error:
        return f"Error: {error}"


@mcp.tool()
def create_spreadsheet(name: str, folder_id: Optional[str] = None) -> str:
    """
    Create a new Google Sheets file, optionally in a specific folder.

    Args:
        name: Name for the new spreadsheet
        folder_id: Optional ID of the folder to create the spreadsheet in

    Returns:
        JSON string with information about the new spreadsheet
    """
    drive_service, sheets_service = get_services()

    try:
        # Create a new blank spreadsheet
        spreadsheet = {"properties": {"title": name}}

        spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet, fields="spreadsheetId").execute()

        # If folder specified, move the file there
        if folder_id:
            # Get the current parents
            file = drive_service.files().get(fileId=spreadsheet["spreadsheetId"], fields="parents").execute()
            previous_parents = ",".join(file.get("parents", []))

            # Move to the specified folder
            file = (
                drive_service.files()
                .update(
                    fileId=spreadsheet["spreadsheetId"],
                    addParents=folder_id,
                    removeParents=previous_parents,
                    fields="id, parents",
                )
                .execute()
            )

        # Get complete file details
        file_details = (
            drive_service.files()
            .get(fileId=spreadsheet["spreadsheetId"], fields="id, name, createdTime, modifiedTime, webViewLink")
            .execute()
        )

        return json.dumps(file_details, indent=2)

    except HttpError as error:
        return f"Error: {error}"


@mcp.tool()
def rename_spreadsheet(file_id: str, new_name: str) -> str:
    """
    Rename a Google Sheets file.

    Args:
        file_id: ID of the spreadsheet to rename
        new_name: New name for the spreadsheet

    Returns:
        JSON string with information about the renamed spreadsheet
    """
    drive_service, _ = get_services()

    try:
        # Update the file metadata
        body = {"name": new_name}
        updated_file = (
            drive_service.files()
            .update(fileId=file_id, body=body, fields="id, name, createdTime, modifiedTime, webViewLink")
            .execute()
        )

        return json.dumps(updated_file, indent=2)

    except HttpError as error:
        return f"Error: {error}"


# ---- Sheet operations ----


@mcp.tool()
def list_sheets(spreadsheet_id: str) -> str:
    """
    List all sheets within a Google Sheets file.

    Args:
        spreadsheet_id: ID of the spreadsheet

    Returns:
        JSON string with sheet information
    """
    _, sheets_service = get_services()

    try:
        # Get spreadsheet information
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        # Extract sheet information
        sheets = spreadsheet.get("sheets", [])
        sheet_details = []

        for sheet in sheets:
            sheet_properties = sheet.get("properties", {})
            sheet_details.append(
                {
                    "sheetId": sheet_properties.get("sheetId"),
                    "title": sheet_properties.get("title"),
                    "index": sheet_properties.get("index"),
                    "sheetType": sheet_properties.get("sheetType"),
                    "rowCount": sheet_properties.get("gridProperties", {}).get("rowCount"),
                    "columnCount": sheet_properties.get("gridProperties", {}).get("columnCount"),
                }
            )

        return json.dumps(sheet_details, indent=2)

    except HttpError as error:
        return f"Error: {error}"


@mcp.tool()
def copy_sheet(spreadsheet_id: str, sheet_id: int, new_sheet_name: str) -> str:
    """
    Copy a sheet within a Google Sheets file.

    Args:
        spreadsheet_id: ID of the spreadsheet
        sheet_id: ID of the sheet to copy
        new_sheet_name: Name for the copied sheet

    Returns:
        JSON string with information about the new sheet
    """
    _, sheets_service = get_services()

    try:
        # Create the copy sheet request
        request = {
            "destination_spreadsheet_id": spreadsheet_id,
            "destination_sheet_id": None,  # This will create a new sheet
        }

        response = (
            sheets_service.spreadsheets()
            .sheets()
            .copyTo(spreadsheetId=spreadsheet_id, sheetId=sheet_id, body=request)
            .execute()
        )

        # Now rename the copied sheet
        requests = [
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": response["sheetId"], "title": new_sheet_name},
                    "fields": "title",
                }
            }
        ]

        result = (
            sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests})
            .execute()
        )

        # Get updated sheet information
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        sheets = spreadsheet.get("sheets", [])
        for sheet in sheets:
            properties = sheet.get("properties", {})
            if properties.get("sheetId") == response["sheetId"]:
                return json.dumps(properties, indent=2)

        return json.dumps(response, indent=2)

    except HttpError as error:
        return f"Error: {error}"


@mcp.tool()
def rename_sheet(spreadsheet_id: str, sheet_id: int, new_name: str) -> str:
    """
    Rename a sheet within a Google Sheets file.

    Args:
        spreadsheet_id: ID of the spreadsheet
        sheet_id: ID of the sheet to rename
        new_name: New name for the sheet

    Returns:
        JSON string with information about the rename operation
    """
    _, sheets_service = get_services()

    try:
        requests = [
            {"updateSheetProperties": {"properties": {"sheetId": sheet_id, "title": new_name}, "fields": "title"}}
        ]

        response = (
            sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests})
            .execute()
        )

        # Return success message with response details
        result = {"success": True, "message": f"Sheet renamed to '{new_name}'", "response": response}

        return json.dumps(result, indent=2)

    except HttpError as error:
        return f"Error: {error}"


# ---- Data operations ----


@mcp.tool()
def get_sheet_data(spreadsheet_id: str, sheet_name: str, range_a1: Optional[str] = None) -> str:
    """
    Get data from a sheet in a Google Sheets file.

    Args:
        spreadsheet_id: ID of the spreadsheet
        sheet_name: Name of the sheet to get data from
        range_a1: Optional A1 notation range (e.g., "A1:D10"). If not provided, gets all data.

    Returns:
        JSON string with the sheet data
    """
    _, sheets_service = get_services()

    try:
        # Form the range string
        if range_a1:
            range_str = f"{sheet_name}!{range_a1}"
        else:
            range_str = sheet_name

        # Get the values
        result = (
            sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=range_str, valueRenderOption="FORMATTED_VALUE")
            .execute()
        )

        values = result.get("values", [])

        # Return the data in a structured format
        return json.dumps({"range": result.get("range"), "values": values}, indent=2)

    except HttpError as error:
        return f"Error: {error}"


@mcp.tool()
def add_rows(spreadsheet_id: str, sheet_id: int, data: List[List[Any]], start_row_index: int = 0) -> str:
    """
    Add rows to a sheet in a Google Sheets file.

    Args:
        spreadsheet_id: ID of the spreadsheet
        sheet_id: ID of the sheet to add rows to
        data: List of rows to add, where each row is a list of values
        start_row_index: Index (0-based) where to insert the rows

    Returns:
        JSON string with information about the operation
    """
    _, sheets_service = get_services()

    try:
        # First get the sheet name from the sheet ID
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        sheet_name = None
        for sheet in spreadsheet.get("sheets", []):
            if sheet.get("properties", {}).get("sheetId") == sheet_id:
                sheet_name = sheet.get("properties", {}).get("title")
                break

        if not sheet_name:
            return json.dumps({"success": False, "message": f"Sheet with ID {sheet_id} not found"}, indent=2)

        # Now add the rows
        range_str = f"{sheet_name}!A{start_row_index + 1}"

        body = {"values": data}

        result = (
            sheets_service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=range_str,
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body=body,
            )
            .execute()
        )

        return json.dumps(
            {
                "success": True,
                "updatedRange": result.get("updates", {}).get("updatedRange"),
                "updatedRows": result.get("updates", {}).get("updatedRows"),
                "updatedColumns": result.get("updates", {}).get("updatedColumns"),
                "updatedCells": result.get("updates", {}).get("updatedCells"),
            },
            indent=2,
        )

    except HttpError as error:
        return f"Error: {error}"


@mcp.tool()
def add_columns(spreadsheet_id: str, sheet_id: int, number_of_columns: int = 1) -> str:
    """
    Add columns to a sheet in a Google Sheets file.

    Args:
        spreadsheet_id: ID of the spreadsheet
        sheet_id: ID of the sheet to add columns to
        number_of_columns: Number of columns to add

    Returns:
        JSON string with information about the operation
    """
    _, sheets_service = get_services()

    try:
        # First get current sheet dimensions
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        current_column_count = 0
        for sheet in spreadsheet.get("sheets", []):
            if sheet.get("properties", {}).get("sheetId") == sheet_id:
                current_column_count = sheet.get("properties", {}).get("gridProperties", {}).get("columnCount", 0)
                break

        if current_column_count == 0:
            return json.dumps({"success": False, "message": f"Sheet with ID {sheet_id} not found"}, indent=2)

        # Now add columns by updating sheet properties
        requests = [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "gridProperties": {"columnCount": current_column_count + number_of_columns},
                    },
                    "fields": "gridProperties.columnCount",
                }
            }
        ]

        response = (
            sheets_service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests})
            .execute()
        )

        return json.dumps(
            {
                "success": True,
                "message": f"{number_of_columns} columns added",
                "previousColumnCount": current_column_count,
                "newColumnCount": current_column_count + number_of_columns,
            },
            indent=2,
        )

    except HttpError as error:
        return f"Error: {error}"


@mcp.tool()
def update_cell(spreadsheet_id: str, sheet_name: str, cell_a1: str, value: Any) -> str:
    """
    Update a single cell in a Google Sheets file.

    Args:
        spreadsheet_id: ID of the spreadsheet
        sheet_name: Name of the sheet containing the cell
        cell_a1: A1 notation for the cell (e.g., "B2")
        value: New value for the cell

    Returns:
        JSON string with information about the update
    """
    _, sheets_service = get_services()

    try:
        range_str = f"{sheet_name}!{cell_a1}"

        body = {"values": [[value]]}  # 2D array for a single cell

        result = (
            sheets_service.spreadsheets()
            .values()
            .update(spreadsheetId=spreadsheet_id, range=range_str, valueInputOption="USER_ENTERED", body=body)
            .execute()
        )

        return json.dumps(
            {"success": True, "updatedRange": result.get("updatedRange"), "updatedCells": result.get("updatedCells")},
            indent=2,
        )

    except HttpError as error:
        return f"Error: {error}"


@mcp.tool()
def update_cells(spreadsheet_id: str, sheet_name: str, range_a1: str, values: List[List[Any]]) -> str:
    """
    Update multiple cells in a Google Sheets file.

    Args:
        spreadsheet_id: ID of the spreadsheet
        sheet_name: Name of the sheet containing the cells
        range_a1: A1 notation for the range (e.g., "A1:C3")
        values: 2D array of values to update the range with

    Returns:
        JSON string with information about the update
    """
    _, sheets_service = get_services()

    try:
        range_str = f"{sheet_name}!{range_a1}"

        body = {"values": values}

        result = (
            sheets_service.spreadsheets()
            .values()
            .update(spreadsheetId=spreadsheet_id, range=range_str, valueInputOption="USER_ENTERED", body=body)
            .execute()
        )

        return json.dumps(
            {
                "success": True,
                "updatedRange": result.get("updatedRange"),
                "updatedRows": result.get("updatedRows"),
                "updatedColumns": result.get("updatedColumns"),
                "updatedCells": result.get("updatedCells"),
            },
            indent=2,
        )

    except HttpError as error:
        return f"Error: {error}"


if __name__ == "__main__":
    mcp.run()
