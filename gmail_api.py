import os
import base64
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime
import time
import re
import requests

class GmailAPI:
    """Class to interact with Gmail API to fetch emails and download attachments."""
    
    def __init__(self, credentials_path='credentials.json', token_path='token.pickle'):
        """
        Initialize the Gmail API client.
        
        Args:
            credentials_path: Path to the Gmail API credentials JSON file
            token_path: Path to save/load the token pickle file
        """
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.scopes = ['https://www.googleapis.com/auth/gmail.readonly']
        self.service = None
        self.authenticate()  # Initialize the service immediately
        
    def authenticate(self):
        """Authenticate with Gmail API and get service object."""
        creds = None
        
        # Load token if it exists
        if os.path.exists(self.token_path):
            with open(self.token_path, 'rb') as token:
                creds = pickle.load(token)
        
        # If there are no (valid) credentials available, let the user log in
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_path, self.scopes)
                creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open(self.token_path, 'wb') as token:
                pickle.dump(creds, token)
        
        # Build the Gmail service
        try:
            self.service = build('gmail', 'v1', credentials=creds)
            return True
        except Exception as e:
            print(f"Error building Gmail service: {e}")
            return False
    
    def get_messages(self, query="from:weatherlink.com", max_results=10):
        """
        Get messages matching a search query.
        
        Args:
            query: Gmail search query
            max_results: Maximum number of results to return
            
        Returns:
            list: List of message objects
        """
        try:
            if not self.service:
                if not self.authenticate():
                    return []
            
            print(f"Searching for messages with query: {query}")
            
            # Get list of messages
            results = self.service.users().messages().list(
                userId='me',
                q=query,
                maxResults=max_results
            ).execute()
            
            messages = results.get('messages', [])
            print(f"Found {len(messages)} messages")
            return messages
            
        except Exception as e:
            print(f"Error getting messages: {e}")
            return []
    
    def get_message_content(self, message_id):
        """
        Get the content of a specific message.
        
        Args:
            message_id: ID of the message to get
            
        Returns:
            dict: Message content object
        """
        try:
            if not self.service:
                if not self.authenticate():
                    return None
            
            message = self.service.users().messages().get(
                userId='me',
                id=message_id,
                format='full'
            ).execute()
            
            return message
            
        except Exception as e:
            print(f"Error getting message content: {e}")
            return None
    
    def download_attachment(self, message_id, attachment_id, attachment_name):
        """
        Download an attachment from a message.
        
        Args:
            message_id: ID of the message containing the attachment
            attachment_id: ID of the attachment to download
            attachment_name: Name to save the attachment as
            
        Returns:
            str: Path to the downloaded attachment or None if download failed
        """
        try:
            if not self.service:
                if not self.authenticate():
                    return None
            
            attachment = self.service.users().messages().attachments().get(
                userId='me',
                messageId=message_id,
                id=attachment_id
            ).execute()
            
            # Decode the attachment data
            file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
            
            # Generate a filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"{attachment_name.split('.')[0]}_{timestamp}.csv"
            
            # Save the attachment
            with open(output_filename, 'wb') as f:
                f.write(file_data)
            
            print(f"Attachment saved as: {output_filename}")
            return output_filename
            
        except Exception as e:
            print(f"Error downloading attachment: {e}")
            return None
    
    def find_weatherlink_export_and_download(self, max_wait_time=300):
        """
        Find the latest WeatherLink export email and download the CSV file from the provided link.
        
        Args:
            max_wait_time: Maximum time to wait for the email (in seconds)
            
        Returns:
            str: Path to the downloaded CSV file or None if not found
        """
        try:
            print("Looking for WeatherLink export emails...")
            
            # Search for messages from WeatherLink
            messages = self.get_messages(
                query="from:weatherlink.com",
                max_results=5
            )
            
            if messages:
                print(f"Found {len(messages)} messages. Processing the latest one...")
                # Get the most recent message (first in the list)
                latest_message = self.get_message_content(messages[0]['id'])
                
                if latest_message and 'payload' in latest_message:
                    # Get the email body
                    if 'parts' in latest_message['payload']:
                        for part in latest_message['payload']['parts']:
                            if part['mimeType'] == 'text/html':
                                # Decode the email body
                                body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                                
                                # Look for the download link
                                download_link_match = re.search(r'href="(https://s3\.amazonaws\.com/export-wl2-live\.weatherlink\.com/data/[^"]+\.csv)"', body)
                                
                                if download_link_match:
                                    download_url = download_link_match.group(1)
                                    print(f"Found download link: {download_url}")
                                    
                                    # Download the file
                                    response = requests.get(download_url)
                                    if response.status_code == 200:
                                        # Generate a filename with timestamp
                                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                        filename = f"weather_data_{timestamp}.csv"
                                        
                                        # Save the file
                                        with open(filename, 'wb') as f:
                                            f.write(response.content)
                                        
                                        print(f"File downloaded successfully: {filename}")
                                        return filename
                                    else:
                                        print(f"Failed to download file. Status code: {response.status_code}")
                                else:
                                    print("No download link found in the email")
                    else:
                        print("No HTML content found in the email")
                else:
                    print("Could not get message content")
            else:
                print("No messages found from WeatherLink")
            
            return None
            
        except Exception as e:
            print(f"Error finding and downloading WeatherLink export: {e}")
            return None 
