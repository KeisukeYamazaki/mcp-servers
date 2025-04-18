import json
import logging
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from mcp.server.fastmcp import Context, FastMCP

# Log configuration
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Set required scopes for Google APIs
SCOPES = ["https://www.googleapis.com/auth/documents", "https://www.googleapis.com/auth/drive"]

# Create MCP server instance
mcp = FastMCP("GoogleDocsServer")


def get_credentials():
    """Get authentication credentials for connecting to Google API"""
    creds = None
    token_path = "token.json"
    credentials_path = "credentials.json"

    # Check if credentials.json exists
    if not os.path.exists(credentials_path):
        logger.error(
            f"Authentication file '{credentials_path}' not found. Please download OAuth client ID from Google Cloud Console."
        )
        raise FileNotFoundError(f"Authentication file '{credentials_path}' not found")

    # token.json stores authenticated user information
    logger.info("Checking authentication credentials...")
    if os.path.exists(token_path):
        logger.info(f"Loading token file '{token_path}'...")
        try:
            with open(token_path, "r") as token_file:
                token_data = json.load(token_file)
                creds = Credentials.from_authorized_user_info(token_data, SCOPES)
                logger.info("Existing credentials loaded")
        except Exception as e:
            logger.error(f"Token file loading error: {e}")
            creds = None

    # Update credentials if expired or non-existent
    if not creds or not creds.valid:
        logger.info("Credentials are invalid or expired. Updating...")
        if creds and creds.expired and creds.refresh_token:
            logger.info("Refreshing token...")
            try:
                creds.refresh(Request())
                logger.info("Credentials updated")
            except Exception as e:
                logger.error(f"Failed to refresh token: {e}")
                creds = None

        if not creds:
            logger.info("Starting new authentication flow. Browser will open...")
            try:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
                creds = flow.run_local_server(port=0)
                logger.info("New credentials acquired")
            except Exception as e:
                logger.error(f"Failed to execute authentication flow: {e}")
                raise

        # Save credentials for next time
        logger.info(f"Saving credentials to '{token_path}'...")
        try:
            with open(token_path, "w") as token:
                token.write(creds.to_json())
            logger.info("Credentials saved")
        except Exception as e:
            logger.error(f"Failed to save token: {e}")

    return creds


def get_folder_id_by_path(drive_service, folder_path):
    """
    Get folder ID from folder path (separated by /)
    Example: "MyDocuments/ProjectA/Reports"
    Create if it doesn't exist
    """
    if not folder_path or folder_path == "/":
        # Root folder case
        return "root"

    parts = folder_path.strip("/").split("/")
    parent_id = "root"

    for part in parts:
        if not part:  # Skip empty parts
            continue

        # Search for folder with current path part
        query = f"name='{part}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"
        results = drive_service.files().list(q=query, spaces="drive", fields="files(id, name)").execute()
        items = results.get("files", [])

        if not items:
            # Create folder if it doesn't exist
            folder_metadata = {"name": part, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
            folder = drive_service.files().create(body=folder_metadata, fields="id").execute()
            parent_id = folder.get("id")
        else:
            # Use existing folder
            parent_id = items[0].get("id")

    return parent_id


@mcp.tool()
def create_document(title: str, content: str = "", folder_path: str = "") -> str:
    """
    Create a new Google Document

    Args:
        title: Document title
        content: Initial document content (optional)
        folder_path: Path to folder where document should be created (example: "MyDocuments/ProjectA")

    Returns:
        ID of the created document
    """
    try:
        # Connect to Google Docs API
        creds = get_credentials()
        docs_service = build("docs", "v1", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)

        # Create empty document
        document = docs_service.documents().create(body={"title": title}).execute()
        document_id = document.get("documentId")

        # Move to specified folder (created in root by default)
        if folder_path:
            # Get folder ID (create if it doesn't exist)
            folder_id = get_folder_id_by_path(drive_service, folder_path)

            # Move file to specified folder
            file = drive_service.files().get(fileId=document_id, fields="parents").execute()
            previous_parents = ",".join(file.get("parents"))

            # Update file to change parent folder
            drive_service.files().update(
                fileId=document_id, addParents=folder_id, removeParents=previous_parents, fields="id, parents"
            ).execute()

            logger.info(f"Document moved to folder '{folder_path}' (ID: {folder_id})")

        # Add content if provided
        if content:
            docs_service.documents().batchUpdate(
                documentId=document_id,
                body={
                    "requests": [
                        {
                            "insertText": {
                                "location": {
                                    "index": 1,
                                },
                                "text": content,
                            }
                        }
                    ]
                },
            ).execute()

        folder_info = f"Folder: {folder_path}" if folder_path else "Root folder"
        return f"Document created. ID: {document_id}, Title: {title}, {folder_info}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def list_folders(parent_folder_path: str = "") -> str:
    """
    Get list of folders in Google Drive

    Args:
        parent_folder_path: Parent folder path (default is root)

    Returns:
        List of folders
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        parent_id = "root"
        if parent_folder_path:
            parent_id = get_folder_id_by_path(drive_service, parent_folder_path)

        # Filter folders only
        query = f"mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"

        # Search folders
        results = drive_service.files().list(q=query, pageSize=50, fields="files(id, name, createdTime)").execute()

        items = results.get("files", [])

        if not items:
            return f"No subfolders found in folder '{parent_folder_path if parent_folder_path else 'Root'}'."

        # Format results
        output = f"Folder list for '{parent_folder_path if parent_folder_path else 'Root'}':\n\n"
        for item in items:
            created_time = datetime.fromisoformat(item["createdTime"].replace("Z", "+00:00"))
            formatted_time = created_time.strftime("%Y-%m-%d %H:%M:%S")
            output += f"Folder name: {item['name']}\nID: {item['id']}\nCreated: {formatted_time}\n\n"

        return output

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def create_folder(folder_path: str) -> str:
    """
    Create a folder in Google Drive

    Args:
        folder_path: Path for the folder to create (example: "MyDocuments/NewFolder")

    Returns:
        Creation result
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get folder ID (create if it doesn't exist)
        folder_id = get_folder_id_by_path(drive_service, folder_path)

        return f"Folder '{folder_path}' created. ID: {folder_id}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def update_document(document_id: str, content: str, append: bool = False) -> str:
    """
    Update an existing Google Document

    Args:
        document_id: ID of the document to update
        content: Content to add or replace
        append: If True, append to existing content; if False, replace

    Returns:
        Update result
    """
    try:
        # Connect to Google Docs API
        creds = get_credentials()
        docs_service = build("docs", "v1", credentials=creds)

        # Get document information
        document = docs_service.documents().get(documentId=document_id).execute()

        # Create update requests
        requests = []

        if append:
            # Add text to the end of document
            end_index = document.get("body").get("content")[-1].get("endIndex")
            requests.append(
                {
                    "insertText": {
                        "location": {
                            "index": end_index - 1,
                        },
                        "text": "\n" + content,
                    }
                }
            )
        else:
            # Replace all document content (delete first then insert)
            requests.extend(
                [
                    {
                        "deleteContentRange": {
                            "range": {
                                "startIndex": 1,
                                "endIndex": document.get("body").get("content")[-1].get("endIndex") - 1,
                            }
                        }
                    },
                    {
                        "insertText": {
                            "location": {
                                "index": 1,
                            },
                            "text": content,
                        }
                    },
                ]
            )

        # Execute batch update
        docs_service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()

        action = "appended" if append else "updated"
        return f"Document {action}. ID: {document_id}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def delete_document(document_id: str) -> str:
    """
    Delete a Google Document (move to trash)

    Args:
        document_id: ID of the document to delete

    Returns:
        Delete operation result
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Delete document (move to trash)
        drive_service.files().delete(fileId=document_id).execute()

        return f"Document deleted. ID: {document_id}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def list_documents(folder_path: str = "", max_results: int = 10) -> str:
    """
    Retrieve a list of Google Documents in the specified folder

    Args:
        folder_path: Folder path (e.g., "MyDocuments/ProjectA")
        max_results: Maximum number of documents to retrieve

    Returns:
        List of documents (ID, title, last modified time)
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get folder ID
        parent_id = "root"
        if folder_path:
            parent_id = get_folder_id_by_path(drive_service, folder_path)

        # Filter only Google Docs files
        query = f"mimeType='application/vnd.google-apps.document' and '{parent_id}' in parents and trashed=false"

        # Search for documents
        results = (
            drive_service.files().list(q=query, pageSize=max_results, fields="files(id, name, modifiedTime)").execute()
        )

        items = results.get("files", [])

        if not items:
            return f"No documents found in folder '{folder_path if folder_path else 'Root'}'."

        # Format results
        output = f"Google Documents in folder '{folder_path if folder_path else 'Root'}':\n\n"
        for item in items:
            modified_time = datetime.fromisoformat(item["modifiedTime"].replace("Z", "+00:00"))
            formatted_time = modified_time.strftime("%Y-%m-%d %H:%M:%S")
            output += f"ID: {item['id']}\nTitle: {item['name']}\nLast Updated: {formatted_time}\n\n"

        return output

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def read_document(document_id: str) -> str:
    """
    Read the contents of a Google Document

    Args:
        document_id: ID of the document to read

    Returns:
        Document contents
    """
    try:
        # Connect to Google Docs API
        creds = get_credentials()
        docs_service = build("docs", "v1", credentials=creds)

        # Retrieve document contents
        document = docs_service.documents().get(documentId=document_id).execute()

        # Extract text content of the document
        title = document.get("title")
        content = ""

        for element in document.get("body").get("content"):
            if "paragraph" in element:
                for para_element in element.get("paragraph").get("elements"):
                    if "textRun" in para_element:
                        content += para_element.get("textRun").get("content")

        return f"Title: {title}\n\n{content}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def move_document(document_id: str, destination_folder_path: str) -> str:
    """
    Move a Google Document to another folder

    Args:
        document_id: ID of the document to move
        destination_folder_path: Path of the destination folder (e.g., "MyDocuments/ProjectB")

    Returns:
        Result of the move operation
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get destination folder ID
        destination_folder_id = get_folder_id_by_path(drive_service, destination_folder_path)

        # Get current parent folders
        file = drive_service.files().get(fileId=document_id, fields="parents").execute()
        previous_parents = ",".join(file.get("parents"))

        # Update file to change parent folders
        drive_service.files().update(
            fileId=document_id, addParents=destination_folder_id, removeParents=previous_parents, fields="id, parents"
        ).execute()

        return f"Moved document (ID: {document_id}) to folder '{destination_folder_path}'."

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def rename_document(document_id: str, new_name: str) -> str:
    """
    Rename a Google Document

    Args:
        document_id: ID of the document to rename
        new_name: New document name

    Returns:
        Result of the rename operation
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get current document information
        file = drive_service.files().get(fileId=document_id, fields="name").execute()
        old_name = file.get("name")

        # Update file name
        drive_service.files().update(fileId=document_id, body={"name": new_name}, fields="id, name").execute()

        return f"Document name changed. ID: {document_id}\nOld Name: {old_name}\nNew Name: {new_name}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def rename_folder(folder_path: str, new_name: str) -> str:
    """
    Rename a Google Drive folder

    Args:
        folder_path: Path of the folder to rename
        new_name: New folder name (not the path, just the folder name itself)

    Returns:
        Result of the rename operation
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get folder ID
        folder_id = get_folder_id_by_path(drive_service, folder_path)

        # Get current folder information
        folder = drive_service.files().get(fileId=folder_id, fields="name").execute()
        old_name = folder.get("name")

        # Update folder name
        drive_service.files().update(fileId=folder_id, body={"name": new_name}, fields="id, name").execute()

        # Get parent folder path (to create new path)
        parent_path = "/".join(folder_path.split("/")[:-1])
        new_path = f"{parent_path}/{new_name}" if parent_path else new_name

        return f"Folder name changed.\nOld Path: {folder_path}\nNew Path: {new_path}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def copy_document(document_id: str, new_name: str = None, destination_folder_path: str = None) -> str:
    """
    Copy a Google Document

    Args:
        document_id: ID of the document to copy
        new_name: New document name (defaults to "Copy of [original name]" if not specified)
        destination_folder_path: Path of the destination folder (defaults to the original folder if not specified)

    Returns:
        Result of the copy operation
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get document information
        file = drive_service.files().get(fileId=document_id, fields="name, parents").execute()
        orig_name = file.get("name")
        orig_parents = file.get("parents", [])

        # Create metadata for copying
        body = {"name": new_name if new_name else f"Copy of {orig_name}"}

        # Copy the document
        copied_file = drive_service.files().copy(fileId=document_id, body=body).execute()
        copied_id = copied_file.get("id")

        # Move to destination folder if specified
        if destination_folder_path:
            destination_folder_id = get_folder_id_by_path(drive_service, destination_folder_path)

            # Get current parent folders
            previous_parents = ",".join(orig_parents)

            # Update file to change parent folders
            drive_service.files().update(
                fileId=copied_id, addParents=destination_folder_id, removeParents=previous_parents, fields="id, parents"
            ).execute()

            location_info = f"Destination Folder: {destination_folder_path}"
        else:
            location_info = "Same location as original folder"

        return f"Document copied.\nOriginal Document: {orig_name} (ID: {document_id})\nNew Document: {copied_file.get('name')} (ID: {copied_id})\n{location_info}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def copy_folder(folder_path: str, new_name: str = None, destination_folder_path: str = None) -> str:
    """
    Copy a Google Drive folder and its contents

    Args:
        folder_path: Path of the folder to copy
        new_name: New folder name (defaults to "Copy of [original name]" if not specified)
        destination_folder_path: Path of the parent folder to copy to (defaults to the same level as the original folder if not specified)

    Returns:
        Result of the copy operation
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get source folder ID
        source_folder_id = get_folder_id_by_path(drive_service, folder_path)

        # Get source folder information
        source_folder = drive_service.files().get(fileId=source_folder_id, fields="name, parents").execute()
        source_name = source_folder.get("name")
        source_parents = source_folder.get("parents", [])

        # Determine new folder name
        new_folder_name = new_name if new_name else f"Copy of {source_name}"

        # Determine destination parent folder ID
        if destination_folder_path:
            parent_folder_id = get_folder_id_by_path(drive_service, destination_folder_path)
        else:
            parent_folder_id = source_parents[0] if source_parents else "root"

        # Create root folder for the copy
        folder_metadata = {
            "name": new_folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_folder_id],
        }

        new_folder = drive_service.files().create(body=folder_metadata, fields="id, name").execute()
        new_folder_id = new_folder.get("id")

        # Recursive function to copy folder contents
        def copy_folder_contents(source_id, target_id):
            # Get files and folders in source folder
            query = f"'{source_id}' in parents and trashed=false"
            results = drive_service.files().list(q=query, fields="files(id, name, mimeType)").execute()
            items = results.get("files", [])

            copied_count = 0

            for item in items:
                item_id = item.get("id")
                item_name = item.get("name")
                item_mime = item.get("mimeType")

                if item_mime == "application/vnd.google-apps.folder":
                    # For folders, create a new folder and recursively copy
                    sub_folder_metadata = {
                        "name": item_name,
                        "mimeType": "application/vnd.google-apps.folder",
                        "parents": [target_id],
                    }
                    sub_folder = drive_service.files().create(body=sub_folder_metadata, fields="id").execute()
                    sub_copied = copy_folder_contents(item_id, sub_folder.get("id"))
                    copied_count += sub_copied + 1  # Count the folder itself
                else:
                    # For files, create a copy
                    file_copy_metadata = {"name": item_name, "parents": [target_id]}
                    drive_service.files().copy(fileId=item_id, body=file_copy_metadata).execute()
                    copied_count += 1

            return copied_count

        # Copy folder contents
        copied_items_count = copy_folder_contents(source_folder_id, new_folder_id)

        # Format destination path
        if destination_folder_path:
            target_path = f"{destination_folder_path}/{new_folder_name}"
        else:
            parent_path = "/".join(folder_path.split("/")[:-1])
            target_path = f"{parent_path}/{new_folder_name}" if parent_path else new_folder_name

        return f"Folder copied.\nOriginal Folder: {folder_path}\nDestination: {target_path}\nNumber of copied items: {copied_items_count}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def lock_document(document_id: str, reason: str = "") -> str:
    """
    Apply editing restrictions to a Google Document (make it read-only)

    Args:
        document_id: ID of the document to lock
        reason: Reason for locking (optional)

    Returns:
        Result of the lock operation
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get document information
        file = drive_service.files().get(fileId=document_id, fields="name, contentRestrictions").execute()
        doc_name = file.get("name")

        # Check current content restrictions
        content_restrictions = file.get("contentRestrictions", [])

        if content_restrictions and content_restrictions[0].get("readOnly"):
            # If already locked
            restricting_user = content_restrictions[0].get("restrictingUser", {}).get("displayName", "Unknown user")
            restriction_date = content_restrictions[0].get("restrictionTime", "Unknown date")
            existing_reason = content_restrictions[0].get("reason", "No reason")

            return (
                f"Document '{doc_name}' (ID: {document_id}) is already locked.\n"
                f"Locked by: {restricting_user}\n"
                f"Lock date: {restriction_date}\n"
                f"Lock reason: {existing_reason}"
            )

        # Set content restrictions on the document
        update_body = {
            "contentRestrictions": [{"readOnly": True, "reason": reason if reason else "Document has been locked"}]
        }

        # Update the file
        updated_file = (
            drive_service.files()
            .update(fileId=document_id, body=update_body, fields="name, contentRestrictions")
            .execute()
        )

        # Get updated information
        updated_restrictions = updated_file.get("contentRestrictions", [])

        if updated_restrictions and updated_restrictions[0].get("readOnly"):
            restriction_info = updated_restrictions[0]
            restricting_user = restriction_info.get("restrictingUser", {}).get("displayName", "Unknown user")
            restriction_date = restriction_info.get("restrictionTime", "Unknown date")

            reason_text = f"\nLock reason: {reason}" if reason else ""

            return (
                f"Document '{doc_name}' (ID: {document_id}) has been locked.\n"
                f"Locked by: {restricting_user}\n"
                f"Lock date: {restriction_date}{reason_text}"
            )
        else:
            return f"Failed to lock document '{doc_name}' (ID: {document_id})."

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def unlock_document(document_id: str) -> str:
    """
    Remove editing restrictions from a Google Document

    Args:
        document_id: ID of the document to unlock

    Returns:
        Result of the unlock operation
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get document information
        file = drive_service.files().get(fileId=document_id, fields="name, contentRestrictions").execute()
        doc_name = file.get("name")

        # Check current content restrictions
        content_restrictions = file.get("contentRestrictions", [])

        if not content_restrictions or not content_restrictions[0].get("readOnly"):
            # If already unlocked
            return f"Document '{doc_name}' (ID: {document_id}) is not locked."

        # Get existing lock information (for logging)
        if content_restrictions:
            restriction_info = content_restrictions[0]
            restricting_user = restriction_info.get("restrictingUser", {}).get("displayName", "Unknown user")
            restriction_date = restriction_info.get("restrictionTime", "Unknown date")
            reason = restriction_info.get("reason", "No reason")

        # Remove content restrictions from the document
        update_body = {"contentRestrictions": [{"readOnly": False}]}

        # Update the file
        updated_file = (
            drive_service.files()
            .update(fileId=document_id, body=update_body, fields="name, contentRestrictions")
            .execute()
        )

        # Check content restrictions after update
        updated_restrictions = updated_file.get("contentRestrictions", [])

        if not updated_restrictions or not updated_restrictions[0].get("readOnly"):
            # Display lock history information
            return (
                f"Document '{doc_name}' (ID: {document_id}) has been unlocked.\n"
                f"Previously locked by: {restricting_user}\n"
                f"Previous lock date: {restriction_date}\n"
                f"Previous lock reason: {reason}"
            )
        else:
            return f"Failed to unlock document '{doc_name}' (ID: {document_id})."

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def add_bulleted_list(document_id: str, items: list, append: bool = True) -> str:
    """
    Add a bulleted list to a Google Document

    Args:
        document_id: ID of the document
        items: List of items for the bulleted list
        append: If True, append to the end of document; if False, replace content

    Returns:
        Result of the operation
    """
    try:
        # Connect to Google Docs API
        creds = get_credentials()
        docs_service = build("docs", "v1", credentials=creds)

        # Get document information
        document = docs_service.documents().get(documentId=document_id).execute()

        requests = []

        if not append:
            # Delete existing content if not appending
            requests.append(
                {
                    "deleteContentRange": {
                        "range": {
                            "startIndex": 1,
                            "endIndex": document.get("body").get("content")[-1].get("endIndex") - 1,
                        }
                    }
                }
            )
            insert_index = 1
        else:
            # Insert at the end of the document
            insert_index = document.get("body").get("content")[-1].get("endIndex") - 1
            # Add a newline before the list if appending
            requests.append(
                {
                    "insertText": {
                        "location": {
                            "index": insert_index,
                        },
                        "text": "\n",
                    }
                }
            )
            insert_index += 1

        # Add each item as a bulleted list item
        for item in items:
            # Insert text for the item
            requests.append(
                {
                    "insertText": {
                        "location": {
                            "index": insert_index,
                        },
                        "text": item + "\n",
                    }
                }
            )

            # Set text as bulleted list
            requests.append(
                {
                    "createParagraphBullets": {
                        "range": {
                            "startIndex": insert_index,
                            "endIndex": insert_index + len(item),
                        },
                        "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE",
                    }
                }
            )

            # Update insert index for next item
            insert_index += len(item) + 1  # +1 for the newline

        # Execute the requests
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests},
        ).execute()

        return f"Bulleted list added to document (ID: {document_id})"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def add_numbered_list(document_id: str, items: list, start_number: int = 1, append: bool = True) -> str:
    """
    Add a numbered list to a Google Document

    Args:
        document_id: ID of the document
        items: List of items for the numbered list
        start_number: Starting number for the list (default is 1)
        append: If True, append to the end of document; if False, replace content

    Returns:
        Result of the operation
    """
    try:
        # Connect to Google Docs API
        creds = get_credentials()
        docs_service = build("docs", "v1", credentials=creds)

        # Get document information
        document = docs_service.documents().get(documentId=document_id).execute()

        requests = []

        if not append:
            # Delete existing content if not appending
            requests.append(
                {
                    "deleteContentRange": {
                        "range": {
                            "startIndex": 1,
                            "endIndex": document.get("body").get("content")[-1].get("endIndex") - 1,
                        }
                    }
                }
            )
            insert_index = 1
        else:
            # Insert at the end of the document
            insert_index = document.get("body").get("content")[-1].get("endIndex") - 1
            # Add a newline before the list if appending
            requests.append(
                {
                    "insertText": {
                        "location": {
                            "index": insert_index,
                        },
                        "text": "\n",
                    }
                }
            )
            insert_index += 1

        # Add each item as a numbered list item
        for item in items:
            # Insert text for the item
            requests.append(
                {
                    "insertText": {
                        "location": {
                            "index": insert_index,
                        },
                        "text": item + "\n",
                    }
                }
            )

            # Set text as numbered list
            requests.append(
                {
                    "createParagraphBullets": {
                        "range": {
                            "startIndex": insert_index,
                            "endIndex": insert_index + len(item),
                        },
                        "bulletPreset": "NUMBERED_DECIMAL_NESTED",
                    }
                }
            )

            # Update insert index for next item
            insert_index += len(item) + 1  # +1 for the newline

        # Execute the requests
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests},
        ).execute()

        return f"Numbered list added to document (ID: {document_id})"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.tool()
def add_table(document_id: str, data: list, append: bool = True) -> str:
    """
    Add a table to a Google Document

    Args:
        document_id: ID of the document
        data: 2D array of table data (first row can be headers)
        append: If True, append to the end of document; if False, replace content

    Returns:
        Result of the operation
    """
    try:
        # Validate data format
        if not data or not isinstance(data, list) or not all(isinstance(row, list) for row in data):
            return "Invalid data format. Please provide a 2D array for the table."

        # Check for consistent row lengths
        row_lengths = [len(row) for row in data]
        if len(set(row_lengths)) > 1:
            return "All rows must have the same number of columns."

        # Connect to Google Docs API
        creds = get_credentials()
        docs_service = build("docs", "v1", credentials=creds)

        # Get document information
        document = docs_service.documents().get(documentId=document_id).execute()

        requests = []

        if not append:
            # Delete existing content if not appending
            requests.append(
                {
                    "deleteContentRange": {
                        "range": {
                            "startIndex": 1,
                            "endIndex": document.get("body").get("content")[-1].get("endIndex") - 1,
                        }
                    }
                }
            )
            insert_index = 1
        else:
            # Insert at the end of the document
            insert_index = document.get("body").get("content")[-1].get("endIndex") - 1
            # Add a newline before the table if appending
            requests.append(
                {
                    "insertText": {
                        "location": {
                            "index": insert_index,
                        },
                        "text": "\n",
                    }
                }
            )
            insert_index += 1

        # Get table dimensions
        rows = len(data)
        cols = len(data[0]) if rows > 0 else 0

        # Create table
        requests.append(
            {
                "insertTable": {
                    "location": {
                        "index": insert_index,
                    },
                    "rows": rows,
                    "columns": cols,
                }
            }
        )

        # Get the document after inserting the table to get table cell locations
        # Execute the table creation request first
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests},
        ).execute()

        # Now get the updated document to find table cells
        updated_document = docs_service.documents().get(documentId=document_id).execute()

        # Find the inserted table
        table_cells = []
        for content in updated_document.get("body").get("content"):
            if "table" in content:
                # We found our table
                for row_index, row in enumerate(content.get("table").get("tableRows")):
                    for col_index, cell in enumerate(row.get("tableCells")):
                        # Get cell's start and end indices
                        start_index = cell.get("content")[0].get("startIndex")
                        end_index = cell.get("content")[0].get("endIndex")

                        # Only add if we have data for this cell
                        if row_index < len(data) and col_index < len(data[row_index]):
                            cell_text = str(data[row_index][col_index])
                            table_cells.append(
                                {
                                    "row": row_index,
                                    "col": col_index,
                                    "start_index": start_index,
                                    "end_index": end_index,
                                    "text": cell_text,
                                }
                            )

        # Create requests to update cell contents
        cell_requests = []
        for cell in table_cells:
            # Delete existing content in the cell
            cell_requests.append(
                {
                    "deleteContentRange": {
                        "range": {
                            "startIndex": cell["start_index"],
                            "endIndex": cell["end_index"] - 1,  # -1 to preserve paragraph mark
                        }
                    }
                }
            )

            # Insert new content
            cell_requests.append(
                {
                    "insertText": {
                        "location": {
                            "index": cell["start_index"],
                        },
                        "text": cell["text"],
                    }
                }
            )

        # Apply cell content updates
        if cell_requests:
            docs_service.documents().batchUpdate(
                documentId=document_id,
                body={"requests": cell_requests},
            ).execute()

        return f"Table added to document (ID: {document_id}) with {rows} rows and {cols} columns"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.resource("gdocs://{document_id}")
def document_resource(document_id: str) -> str:
    """
    Provides a Google Document with the specified ID as a resource

    Args:
        document_id: ID of the document

    Returns:
        Content of the document
    """
    try:
        # Connect to Google Docs API
        creds = get_credentials()
        docs_service = build("docs", "v1", credentials=creds)

        # Get document content
        document = docs_service.documents().get(documentId=document_id).execute()

        # Extract document text content
        title = document.get("title")
        content = ""

        for element in document.get("body").get("content"):
            if "paragraph" in element:
                for para_element in element.get("paragraph").get("elements"):
                    if "textRun" in para_element:
                        content += para_element.get("textRun").get("content")

        return f"# {title}\n\n{content}"

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.resource("gdocs://folder/{folder_path}")
def folder_documents_resource(folder_path: str) -> str:
    """
    Provides a list of Google Documents in the specified folder as a resource

    Args:
        folder_path: Path to the folder

    Returns:
        List of documents
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Get folder ID
        parent_id = get_folder_id_by_path(drive_service, folder_path)

        # Filter for Google Docs files only
        query = f"mimeType='application/vnd.google-apps.document' and '{parent_id}' in parents and trashed=false"

        # Search for documents
        results = drive_service.files().list(q=query, pageSize=50, fields="files(id, name, modifiedTime)").execute()

        items = results.get("files", [])

        if not items:
            return f"# Documents in folder '{folder_path}'\n\nNo documents found."

        # Format results
        output = f"# Documents in folder '{folder_path}'\n\n"
        for item in items:
            modified_time = datetime.fromisoformat(item["modifiedTime"].replace("Z", "+00:00"))
            formatted_time = modified_time.strftime("%Y-%m-%d %H:%M:%S")
            output += f"- **{item['name']}**\n  - ID: `{item['id']}`\n  - Last modified: {formatted_time}\n\n"

        return output

    except HttpError as error:
        return f"An error occurred: {error}"


@mcp.resource("gdocs://list")
def documents_list_resource() -> str:
    """
    Provides a list of the user's Google Documents as a resource

    Returns:
        List of documents
    """
    try:
        # Connect to Google Drive API
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        # Filter for Google Docs files only
        query = "mimeType='application/vnd.google-apps.document'"

        # Search for documents
        results = drive_service.files().list(q=query, pageSize=50, fields="files(id, name, modifiedTime)").execute()

        items = results.get("files", [])

        if not items:
            return "# Google Documents List\n\nNo documents found."

        # Format results
        output = "# Google Documents List\n\n"
        for item in items:
            modified_time = datetime.fromisoformat(item["modifiedTime"].replace("Z", "+00:00"))
            formatted_time = modified_time.strftime("%Y-%m-%d %H:%M:%S")
            output += f"- **{item['name']}**\n  - ID: `{item['id']}`\n  - Last modified: {formatted_time}\n\n"

        return output

    except HttpError as error:
        return f"An error occurred: {error}"


if __name__ == "__main__":
    # Test authentication at startup
    try:
        logger.info("Testing Google authentication...")
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)
        docs_service = build("docs", "v1", credentials=creds)
        logger.info("Google authentication successful!")

        # Start the server
        logger.info("Starting MCP server...")
        mcp.run()
    except Exception as e:
        logger.error(f"Startup error: {e}")
        import traceback

        traceback.print_exc()
