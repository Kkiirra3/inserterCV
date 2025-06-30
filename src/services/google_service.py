import os
import io
import pickle
import re
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from docx import Document

class GoogleServiceManager:
    SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']

    def get_credentials(self):
        """
        Gets or refreshes credentials for accessing Google API.
        """
        creds = None
        if os.path.exists('creds/token.pickle'):
            with open('creds/token.pickle', 'rb') as token:
                creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('creds/credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open('creds/token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        return creds

    def get_drive_service(self):
        """Gets authenticated Google Drive service"""
        creds = self.get_credentials()
        return build('drive', 'v3', credentials=creds)

    def get_document_id_from_url(self, url):
        """
        Extracts document ID from Google Docs URL.
        """
        pattern = r'/d/([a-zA-Z0-9-_]+)'
        match = re.search(pattern, url)
        if match:
            return match.group(1)
        raise ValueError("Invalid Google Docs URL format")

    def export_to_docx(self, service, doc_id, output_path):
        """
        Exports Google Doc to .docx format
        """
        try:
            request = service.files().export_media(
                fileId=doc_id,
                mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            
            while not done:
                status, done = downloader.next_chunk()
            
            fh.seek(0)
            with open(output_path, 'wb') as f:
                f.write(fh.read())
                f.close()
            
            # Check file contents
            doc = Document(output_path)
            return True
        
        except Exception as e:
            print(f"An error occurred while exporting: {str(e)}")
            return False

    def upload_to_drive(self, service, file_path, title, bullet_color=None):
        """
        Uploads file to Google Drive and converts it to Google Docs
        """
        try:
            file_metadata = {
                'name': title,
                'mimeType': 'application/vnd.google-apps.document'
            }
            
            media = MediaFileUpload(
                file_path,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                resumable=True
            )
            
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            
            doc_id = file.get('id')
            if doc_id:
                # Create service for Google Docs API
                docs_service = build('docs', 'v1', credentials=service._http.credentials)
                
                # Get document for analysis
                document = docs_service.documents().get(documentId=doc_id).execute()
                
                # Collect all update requests
                requests = []
                
                def process_structural_elements(elements, in_table=False):
                    """Recursively processes document's structural elements"""
                    in_responsibilities = False
                    
                    for element in elements:
                        if 'paragraph' in element:
                            # Check only paragraphs inside tables
                            if not in_table:
                                continue
                                
                            para = element['paragraph']
                            text = ''
                            # Collect all paragraph text
                            for section in para.get('elements', []):
                                if 'textRun' in section:
                                    text += section['textRun'].get('content', '')
                            
                            text = text.strip()
                            if not text:
                                continue
                                
                            # Determine if we're in Responsibilities section
                            if text.startswith('Responsibilities'):
                                in_responsibilities = True
                                continue
                            elif text.startswith(('Project roles', 'Period', 'Environment')):
                                in_responsibilities = False
                                continue
                                
                            # Add bullet points only for elements in Responsibilities section
                            if in_responsibilities:
                                start_index = element.get('startIndex')
                                end_index = element.get('endIndex')
                                
                                if start_index is not None and end_index is not None and start_index < end_index:
                                    # 1) create bullets
                                    requests.append({
                                        'createParagraphBullets': {
                                            'range': {'startIndex': start_index,
                                                      'endIndex': end_index},
                                            'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                                        }
                                    })

                                    if bullet_color:
                                        r, g, b = (int(bullet_color[i:i+2], 16) / 255.0
                                                   for i in (0, 2, 4))

                                        # 2) color TAB symbol ⇒ bullet point will be colored too
                                        requests.append({
                                            'updateTextStyle': {
                                                'range': {'startIndex': start_index,
                                                          'endIndex': start_index + 1},
                                                'textStyle': {
                                                    'foregroundColor': {
                                                        'color': {'rgbColor': {'red': r, 'green': g, 'blue': b}}
                                                    }
                                                },
                                                'fields': 'foregroundColor'
                                            }
                                        })

                                        # 3) (optional) return main text to black
                                        requests.append({
                                            'updateTextStyle': {
                                                'range': {'startIndex': start_index + 1,
                                                          'endIndex': end_index},
                                                'textStyle': {
                                                    'foregroundColor': {
                                                        'color': {'rgbColor': {'red': 0, 'green': 0, 'blue': 0}}
                                                    }
                                                },
                                                'fields': 'foregroundColor'
                                            }
                                        })
                        
                        elif 'table' in element:
                            # Process table
                            for row in element['table'].get('tableRows', []):
                                for cell in row.get('tableCells', []):
                                    process_structural_elements(cell.get('content', []), True)
                        
                        elif 'tableOfContents' in element:
                            # Process table of contents
                            process_structural_elements(element['tableOfContents'].get('content', []), in_table)
                
                # Process document content
                process_structural_elements(document.get('body', {}).get('content', []))
                
                # If there are update requests, send them
                if requests:
                    try:
                        for idx, req in enumerate(requests):
                            if 'updateTextStyle' in req:
                                style = req['updateTextStyle']
                                color = style['textStyle']['foregroundColor']['color']['rgbColor']
                        
                        result = docs_service.documents().batchUpdate(
                            documentId=doc_id,
                            body={'requests': requests}
                        ).execute()
                    except Exception as e:
                        print(f"Warning: Failed to apply updates: {str(e)}")
                        print(f"Request that failed: {requests[-1]}")
                
                return doc_id
            
            return file.get('id')
        
        except Exception as e:
            print(f"An error occurred while uploading: {str(e)}")
            return None 