"""
Email processor for cleaning and preparing email data for AI processing.
Handles metadata removal and preparation for AI inputs.
"""

import re
from typing import Dict, Any, List
from bs4 import BeautifulSoup

class EmailProcessor:
    """Process and clean email content from Graph API"""
    
    @staticmethod
    def remove_metadata(email_content: str) -> str:
        """
        Remove metadata and formatting from email content
        Returns only the clean email body text
        """
        if not email_content:
            return ""
            
        # Check if content is HTML
        is_html = bool(re.search(r'<html|<body|<div|<span|<p|<table', email_content, re.IGNORECASE))
        
        # Process HTML content
        if is_html:
            try:
                soup = BeautifulSoup(email_content, 'html.parser')
                
                # Remove script and style elements
                for element in soup(['script', 'style']):
                    element.decompose()
                
                # Try to find the main content area
                main_content = None
                
                # Look for common email content containers
                for div in soup.find_all(['div', 'td']):
                    # Look for content divs with specific classes or IDs
                    if div.get('class') and any(c.lower() in ['content', 'body', 'message', 'email'] 
                                              for c in div.get('class') if c):
                        main_content = div
                        break
                    if div.get('id') and any(id_val.lower() in ['content', 'body', 'message', 'email'] 
                                           for id_val in [div.get('id')] if div.get('id')):
                        main_content = div
                        break
                
                # If we found specific content container, use it; otherwise use body or entire HTML
                if main_content:
                    content = main_content.get_text(separator=' ', strip=True)
                else:
                    # Try to get body if it exists
                    body = soup.find('body')
                    if body:
                        content = body.get_text(separator=' ', strip=True)
                    else:
                        content = soup.get_text(separator=' ', strip=True)
                        
            except Exception as e:

                print(f"Error parsing HTML: {e}")

                # Fallback to simple HTML tag removal
            content = re.sub(r'<.*?>', ' ', email_content)

        else:
            content = email_content

        # Process the content line by line
        lines = content.split('\n')
        cleaned_lines = []
        content_started = False
        skip_section = False
        
        # Enhanced metadata markers to identify lines to skip
        metadata_markers = [
            "From:", "To:", "Sent:", "Date:", "Subject:",
            "CAUTION:", "Original Message", "Forwarded Message",
            "------ Original Message ------", "This message and any attachments",
            "This email and any files", "DISCLAIMER", "Confidential",
            "This is a private communication", "NOTICE:",
            "CONFIDENTIALITY NOTICE", "PRIVILEGED AND CONFIDENTIAL",
            "Legal Notice:", "Sent from my iPhone", "Sent from my mobile device",
            "----------------------", "-----Original Message-----",
            "Begin forwarded message", "On ", "wrote:",
            
        ]
        
        # Process each line
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Check if this line starts a section to skip
            if any(marker in line for marker in metadata_markers):
                skip_section = True
                continue
            
            # If we find markers that likely indicate the start of actual content
            if any(marker in line for marker in ["Dear ", "Hi ", "Hello ", "Good morning", "Good afternoon"]):
                content_started = True
                skip_section = False
            
            # Add content lines that aren't in skip sections
            if not skip_section and (content_started or len(cleaned_lines) == 0):
                cleaned_lines.append(line)
        
        # Join lines and clean up extra whitespace
        cleaned_content = ' '.join(cleaned_lines)
        
        # Remove excessive whitespace and normalize spaces
        cleaned_content = re.sub(r'\s+', ' ', cleaned_content).strip()
        
        # Preserve property details by ensuring addresses and dates are kept intact
        property_patterns = [
            r'\b\d+\s+[A-Za-z\s]+(?:Road|Rd|Street|St|Avenue|Ave|Boulevard|Blvd|Drive|Dr|Lane|Ln|Way|Place|Pl|Court|Ct|Terrace|Ter|Trail|Trl|Park|Circle|Cir)\b',
            r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?,\s+\d{4}\b',
            r'\b\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)\b'
        ]
        
        # Extract property details from original content
        property_details = []
        for pattern in property_patterns:
            matches = re.finditer(pattern, content, re.IGNORECASE)
            for match in matches:
                property_details.append(match.group(0))
        
        # Make sure we preserve property details
        for detail in property_details:
            if detail not in cleaned_content:
                cleaned_content += f" {detail}"
        
        return cleaned_content.strip()
    
    @staticmethod
    def format_for_ai(email: Dict[str, Any]) -> Dict[str, Any]:
        """
        Format email data into a structure suitable for AI processing
        Returns a dictionary with cleaned and formatted email fields
        """
        # Extract and clean main content
        body = email.get('body', {})
        content = ""
        
        if isinstance(body, dict):
            content = body.get('content', '')
        elif isinstance(body, str):
            content = body
            
        clean_content = EmailProcessor.remove_metadata(content)
        
        # Extract sender information
        sender = email.get('sender', {}) or email.get('from', {})
        sender_email = ""
        if isinstance(sender, dict):
            sender_email = sender.get('emailAddress', {}).get('address', '')
        else:
            sender_email = str(sender)
            
        # Extract recipient information
        recipients = email.get('toRecipients', [])
        recipient_emails = []
        for recipient in recipients:
            if isinstance(recipient, dict):
                email_addr = recipient.get('emailAddress', {}).get('address', '')
                if email_addr:
                    recipient_emails.append(email_addr)
        
        # Format the result
        return {
            'from_email': sender_email,
            'to_email': recipient_emails,
            'datetime': email.get('receivedDateTime', ''),
            'subject': email.get('subject', ''),
            'content': clean_content,
            'raw_email': email  # Keep the original email for reference
        }
    
    @staticmethod
    def batch_process(emails: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Process a batch of emails for AI analysis"""
        return [EmailProcessor.format_for_ai(email) for email in emails]