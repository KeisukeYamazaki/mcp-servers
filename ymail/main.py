# yahoo_mail_server.py
from mcp.server.fastmcp import FastMCP, Context
import imaplib
import email
from email.header import decode_header
import os
from typing import List, Dict, Any, Optional
import re
from datetime import datetime
import json
from dataclasses import dataclass
import traceback
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator
import base64
import binascii


@asynccontextmanager
async def server_lifespan(server: FastMCP) -> AsyncIterator[dict]:
    """Manage server startup and shutdown lifecycle."""
    try:
        yield {}
    finally:
        if email_config and email_config.connection:
            try:
                email_config.connection.logout()
            except:
                pass


# Create an MCP server with lifespan support
mcp = FastMCP("Yahoo Mail", lifespan=server_lifespan)


# Email connection class
@dataclass
class EmailConfig:
    username: str
    password: str
    imap_server: str = "imap.mail.yahoo.co.jp"
    imap_port: int = 993
    connection: Optional[imaplib.IMAP4_SSL] = None


# Global connection state
email_config = None


# Helper functions for IMAP operations
def connect_to_email(ctx: Context) -> bool:
    """Establish connection to Yahoo Mail via IMAP"""
    global email_config
    
    if email_config and email_config.connection:
        try:
            # Check if connection is still alive
            status, _ = email_config.connection.noop()
            if status == 'OK':
                return True
        except:
            # Connection is dead, we'll reconnect
            pass
    
    try:
        # Get credentials from environment or context
        username = os.environ.get("YAHOO_EMAIL")
        password = os.environ.get("YAHOO_PASSWORD")
        
        if not username or not password:
            ctx.error("Email credentials not found. Set YAHOO_EMAIL and YAHOO_PASSWORD environment variables.")
            return False
        
        # Create new connection
        connection = imaplib.IMAP4_SSL(email_config.imap_server if email_config else "imap.mail.yahoo.co.jp", 
                                       email_config.imap_port if email_config else 993)
        connection.login(username, password)
        
        # Store connection
        email_config = EmailConfig(
            username=username,
            password=password,
            connection=connection
        )
        
        ctx.info(f"Successfully connected to {username}")
        return True
    except Exception as e:
        ctx.error(f"Failed to connect to Yahoo Mail: {str(e)}")
        return False


def decode_email_header(header):
    """Decode email header"""
    decoded_headers = decode_header(header)
    header_text = ""
    
    for value, encoding in decoded_headers:
        if isinstance(value, bytes):
            try:
                if encoding:
                    header_text += value.decode(encoding)
                else:
                    header_text += value.decode('utf-8', errors='replace')
            except:
                header_text += value.decode('utf-8', errors='replace')
        else:
            header_text += str(value)
            
    return header_text


def get_email_body(msg):
    """Extract email body preferring text/plain"""
    if msg.is_multipart():
        text_content = ""
        html_content = ""
        
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            
            # Skip attachments
            if "attachment" in content_disposition:
                continue
                
            # Get the body
            try:
                body = part.get_payload(decode=True)
                
                if body:
                    if content_type == "text/plain":
                        text_content = body.decode('utf-8', errors='replace')
                    elif content_type == "text/html":
                        html_content = body.decode('utf-8', errors='replace')
            except:
                pass
        
        # Prefer plain text if available
        return text_content if text_content else html_content
    else:
        # Not multipart
        try:
            return msg.get_payload(decode=True).decode('utf-8', errors='replace')
        except:
            return "Could not decode email body"


# More accurate IMAP modified UTF-7 decoder
def decode_modified_utf7(s):
    """
    Decode IMAP's modified UTF-7 encoding for folder names.
    IMAP's UTF-7 is a variant of UTF-7 as defined in RFC 3501.
    """
    if not s:
        return s

    result = ""
    i = 0
    while i < len(s):
        if s[i] == '&' and i + 1 < len(s):
            # Find the end of the Base64 encoded sequence
            j = s.find('-', i + 1)
            if j == -1:
                # No closing dash, treat as literal '&'
                result += '&'
                i += 1
            elif j == i + 1:
                # '&-' means literal '&'
                result += '&'
                i += 2
            else:
                # Extract the Base64 encoded sequence
                encoded = s[i+1:j].replace(',', '/')
                
                try:
                    # Add proper Base64 padding if necessary
                    padding = len(encoded) % 4
                    if padding > 0:
                        encoded += '=' * (4 - padding)
                    
                    # Base64 decode
                    decoded_bytes = base64.b64decode(encoded)
                    
                    # Handle UTF-16BE encoding (IMAP UTF-7 variant uses BE format)
                    # This is the key part: IMAP modified UTF-7 encodes directly to UTF-16BE
                    decoded_text = decoded_bytes.decode('utf-16be')
                    result += decoded_text
                except (UnicodeDecodeError, binascii.Error):
                    # Fallback if decoding fails
                    result += '&' + s[i+1:j] + '-'
                    
                i = j + 1
        else:
            # For standard ASCII characters
            result += s[i]
            i += 1
            
    return result


# MCP Tools
@mcp.tool()
def list_folders(ctx: Context) -> str:
    """List all available email folders"""
    if not connect_to_email(ctx):
        return "Failed to connect to Yahoo Mail"
    
    try:
        status, folders = email_config.connection.list()
        
        if status != 'OK':
            return f"Failed to retrieve folders: {status}"
        
        folder_list = []
        for folder in folders:
            try:
                # Parse the folder line: e.g., '(\HasNoChildren) "/" "INBOX"'
                folder_str = folder.decode('utf-8', errors='replace')
                
                # Debug information to see what we're getting
                ctx.info(f"Raw folder: {folder_str}")
                
                # Extract the folder name (last quoted part)
                match = re.search(r'"([^"]+)"$', folder_str)
                if match:
                    folder_name = match.group(1)
                    # Log before decoding
                    ctx.info(f"Before decoding: {folder_name}")
                    
                    # Decode from IMAP's modified UTF-7
                    decoded_name = decode_modified_utf7(folder_name)
                    
                    # Log after decoding
                    ctx.info(f"After decoding: {decoded_name}")
                    
                    folder_list.append(decoded_name)
                else:
                    # Fallback if pattern doesn't match
                    folder_parts = folder_str.split(' "')
                    if len(folder_parts) >= 2:
                        folder_name = folder_parts[-1].strip('"')
                        decoded_name = decode_modified_utf7(folder_name)
                        folder_list.append(decoded_name)
                    else:
                        ctx.error(f"Could not parse folder structure: {folder_str}")
            except Exception as e:
                ctx.error(f"Error processing folder: {str(e)}\n{traceback.format_exc()}")
                continue
        
        # Log the final folder list for debugging
        ctx.info(f"Final folder list: {folder_list}")
        
        return json.dumps({"folders": folder_list}, indent=2, ensure_ascii=False)
    except Exception as e:
        ctx.error(f"Error listing folders: {str(e)}\n{traceback.format_exc()}")
        return f"Error listing folders: {str(e)}"


@mcp.tool()
def search_emails(query: str, folder: str = "INBOX", limit: int = 10, ctx: Context = None) -> str:
    """
    Search for emails in a specified folder
    
    Args:
        query: Search query (e.g., "FROM user@example.com", "SUBJECT meeting", "TEXT important", "SINCE 01-Jan-2023")
        folder: Email folder to search in (default: INBOX)
        limit: Maximum number of emails to return (default: 10)
    """
    if not connect_to_email(ctx):
        return "Failed to connect to Yahoo Mail"
    
    try:
        # Select the folder
        status, _ = email_config.connection.select(folder)
        if status != 'OK':
            return f"Failed to select folder {folder}"
        
        # Parse the query to match IMAP syntax
        imap_query = query
        if not any(x in query.upper() for x in ['FROM', 'TO', 'SUBJECT', 'BODY', 'TEXT', 'SINCE', 'BEFORE', 'ON']):
            # Default to ALL if no specific query is provided
            imap_query = 'ALL'
        
        # Search for emails
        status, messages = email_config.connection.search(None, imap_query)
        if status != 'OK':
            return f"Search failed: {status}"
        
        # Get message IDs
        message_ids = messages[0].split()
        
        # Limit the number of results
        message_ids = message_ids[-min(limit, len(message_ids)):]
        
        results = []
        for msg_id in message_ids:
            status, msg_data = email_config.connection.fetch(msg_id, '(RFC822)')
            if status != 'OK':
                continue
                
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # Parse email details
            subject = decode_email_header(msg.get('Subject', 'No Subject'))
            from_addr = decode_email_header(msg.get('From', 'Unknown'))
            date = decode_email_header(msg.get('Date', 'Unknown'))
            
            results.append({
                "id": msg_id.decode(),
                "subject": subject,
                "from": from_addr,
                "date": date
            })
        
        if not results:
            return f"No emails found matching query: {query}"
            
        return json.dumps({"results": results, "count": len(results), "folder": folder}, indent=2, ensure_ascii=False)
    except Exception as e:
        ctx.error(f"Error searching emails: {str(e)}\n{traceback.format_exc()}")
        return f"Error searching emails: {str(e)}"


@mcp.tool()
def read_email(email_id: str, folder: str = "INBOX", ctx: Context = None) -> str:
    """
    Read a specific email by ID
    
    Args:
        email_id: Email ID to fetch
        folder: Folder where the email is located (default: INBOX)
    """
    if not connect_to_email(ctx):
        return "Failed to connect to Yahoo Mail"
    
    try:
        # Select the folder
        status, _ = email_config.connection.select(folder)
        if status != 'OK':
            return f"Failed to select folder {folder}"
        
        # Fetch the email
        status, msg_data = email_config.connection.fetch(email_id.encode(), '(RFC822)')
        if status != 'OK':
            return f"Failed to fetch email with ID {email_id}"
        
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)
        
        # Parse email details
        subject = decode_email_header(msg.get('Subject', 'No Subject'))
        from_addr = decode_email_header(msg.get('From', 'Unknown'))
        to_addr = decode_email_header(msg.get('To', 'Unknown'))
        date = decode_email_header(msg.get('Date', 'Unknown'))
        body = get_email_body(msg)
        
        email_data = {
            "id": email_id,
            "subject": subject,
            "from": from_addr,
            "to": to_addr,
            "date": date,
            "body": body[:2000] + ("..." if len(body) > 2000 else "")  # Truncate long bodies
        }
        
        return json.dumps(email_data, indent=2, ensure_ascii=False)
    except Exception as e:
        ctx.error(f"Error reading email: {str(e)}\n{traceback.format_exc()}")
        return f"Error reading email: {str(e)}"


@mcp.tool()
def get_unread_count(folder: str = "INBOX", ctx: Context = None) -> str:
    """Get the number of unread emails in a folder"""
    if not connect_to_email(ctx):
        return "Failed to connect to Yahoo Mail"
    
    try:
        # Select the folder
        status, data = email_config.connection.select(folder)
        if status != 'OK':
            return f"Failed to select folder {folder}"
        
        # Search for unread emails
        status, messages = email_config.connection.search(None, 'UNSEEN')
        if status != 'OK':
            return f"Failed to search for unread emails: {status}"
        
        # Count unread messages
        unread_count = len(messages[0].split())
        
        return json.dumps({"folder": folder, "unread_count": unread_count}, ensure_ascii=False)
    except Exception as e:
        ctx.error(f"Error getting unread count: {str(e)}\n{traceback.format_exc()}")
        return f"Error getting unread count: {str(e)}"


@mcp.tool()
def mark_as_read(email_id: str, folder: str = "INBOX", ctx: Context = None) -> str:
    """Mark an email as read"""
    if not connect_to_email(ctx):
        return "Failed to connect to Yahoo Mail"
    
    try:
        # Select the folder
        status, _ = email_config.connection.select(folder)
        if status != 'OK':
            return f"Failed to select folder {folder}"
        
        # Mark as read by adding the \\Seen flag
        status, data = email_config.connection.store(email_id.encode(), '+FLAGS', '\\Seen')
        if status != 'OK':
            return f"Failed to mark email as read: {status}"
        
        return f"Email {email_id} marked as read"
    except Exception as e:
        ctx.error(f"Error marking email as read: {str(e)}\n{traceback.format_exc()}")
        return f"Error marking email as read: {str(e)}"


@mcp.tool()
def mark_as_unread(email_id: str, folder: str = "INBOX", ctx: Context = None) -> str:
    """Mark an email as unread"""
    if not connect_to_email(ctx):
        return "Failed to connect to Yahoo Mail"
    
    try:
        # Select the folder
        status, _ = email_config.connection.select(folder)
        if status != 'OK':
            return f"Failed to select folder {folder}"
        
        # Mark as unread by removing the \\Seen flag
        status, data = email_config.connection.store(email_id.encode(), '-FLAGS', '\\Seen')
        if status != 'OK':
            return f"Failed to mark email as unread: {status}"
        
        return f"Email {email_id} marked as unread"
    except Exception as e:
        ctx.error(f"Error marking email as unread: {str(e)}\n{traceback.format_exc()}")
        return f"Error marking email as unread: {str(e)}"


@mcp.tool()
def move_email(email_id: str, source_folder: str, destination_folder: str, ctx: Context = None) -> str:
    """Move an email from one folder to another"""
    if not connect_to_email(ctx):
        return "Failed to connect to Yahoo Mail"
    
    try:
        # Select the source folder
        status, _ = email_config.connection.select(source_folder)
        if status != 'OK':
            return f"Failed to select folder {source_folder}"
        
        # Copy the email to the destination folder
        status, data = email_config.connection.copy(email_id.encode(), destination_folder)
        if status != 'OK':
            return f"Failed to copy email to {destination_folder}: {status}"
        
        # Delete the email from the source folder
        status, data = email_config.connection.store(email_id.encode(), '+FLAGS', '\\Deleted')
        if status != 'OK':
            return f"Failed to mark email for deletion: {status}"
        
        # Expunge to actually delete the message
        email_config.connection.expunge()
        
        return f"Email {email_id} moved from {source_folder} to {destination_folder}"
    except Exception as e:
        ctx.error(f"Error moving email: {str(e)}\n{traceback.format_exc()}")
        return f"Error moving email: {str(e)}"


@mcp.tool()
def delete_email(email_id: str, folder: str = "INBOX", ctx: Context = None) -> str:
    """Delete an email by moving it to Trash"""
    if not connect_to_email(ctx):
        return "Failed to connect to Yahoo Mail"
    
    try:
        # Select the folder
        status, _ = email_config.connection.select(folder)
        if status != 'OK':
            return f"Failed to select folder {folder}"
        
        # Move to Trash by adding the \\Deleted flag
        status, data = email_config.connection.store(email_id.encode(), '+FLAGS', '\\Deleted')
        if status != 'OK':
            return f"Failed to mark email for deletion: {status}"
        
        # Expunge to actually delete the message
        email_config.connection.expunge()
        
        return f"Email {email_id} deleted"
    except Exception as e:
        ctx.error(f"Error deleting email: {str(e)}\n{traceback.format_exc()}")
        return f"Error deleting email: {str(e)}"


# Run the server if executed directly
if __name__ == "__main__":
    mcp.run()
