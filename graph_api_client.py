"""
Microsoft Graph API client for fetching emails from Outlook.
"""

import json
import requests
import re
from datetime import datetime
from typing import Dict, Any, List, Optional
from msal import ConfidentialClientApplication
from thefuzz import fuzz

class GraphAPIClient:
    """Client for Microsoft Graph API to fetch emails from Outlook"""
    
    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
    
    def get_access_token(self) -> Optional[str]:
        """Get an access token for Microsoft Graph API"""
        try:
            msal_authority = f"https://login.microsoftonline.com/{self.tenant_id}"
            msal_scope = ["https://graph.microsoft.com/.default"]
            
            msal_app = ConfidentialClientApplication(
                client_id=self.client_id, 
                client_credential=self.client_secret, 
                authority=msal_authority
            )
            
            result = msal_app.acquire_token_silent(scopes=msal_scope, account=None)
            
            if not result:
                result = msal_app.acquire_token_for_client(scopes=msal_scope)
                
            if "access_token" in result:
                self.access_token = result["access_token"]
                return self.access_token
            else:
                print(f"Error getting token: {result.get('error')}, {result.get('error_description')}")
                return None
                
        except Exception as e:
            print(f"Error getting access token: {str(e)}")
            return None
    
    def get_recent_emails(self, user_email: str, count: int = 20, 
                         filters: Optional[Dict[str, str]] = None) -> List[Dict[str, Any]]:
        """
        Get recent emails for a user with optional filters
        Returns a list of email data from Graph API
        """
        if not self.access_token:
            self.access_token = self.get_access_token()
            
        if not self.access_token:
            print("No access token available. Authentication failed.")
            return []
            
        print(f"Fetching emails with filters: {filters}")
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        base_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        
        # Build query parameters
        params = {
            "$select": "id,subject,sender,receivedDateTime,conversationId",
            "$top": count,
            "$orderby": "receivedDateTime desc"
        }
        
        # Handle filters if provided
        if filters:
            # Special case: If we have a sender filter, use the optimized sender search function
            if "from" in filters and filters["from"]:
                print(f"Using optimized sender search for: {filters['from']}")
                sender = filters["from"]
                # Extract other filters to apply after sender search
                other_filters = {k: v for k, v in filters.items() if k != "from"}
                return self._search_by_sender(user_email, sender, count, other_filters)
            
            # For other filter combinations
            filter_parts = []
            
            # Subject filter
            if "subject" in filters and filters["subject"]:
                filter_parts.append(f"contains(subject,'{filters['subject']}')")
                
            # Handle date filters (these are generally safe to use directly)
            if "date_from" in filters and filters["date_from"]:
                filter_parts.append(f"receivedDateTime ge {filters['date_from']}T00:00:00Z")
                
            if "date_to" in filters and filters["date_to"]:
                filter_parts.append(f"receivedDateTime le {filters['date_to']}T23:59:59Z")
                
            # Recipient filter (can cause complexity issues, handle carefully)
            if "to" in filters and filters["to"]:
                # Try a simpler approach that might work better
                filter_parts.append(f"toRecipients/any(r:contains(r/emailAddress/address,'{filters['to']}'))")
            
            # Add filter parameter if we have any filter parts
            if filter_parts:
                params["$filter"] = " and ".join(filter_parts)
                print(f"Filter query: {params['$filter']}")
        
        try:
            print(f"Making API request to: {base_url}")
            print(f"Query parameters: {params}")
            response = requests.get(base_url, headers=headers, params=params)
            
            if response.status_code == 200:
                result = response.json()
                emails = [self._format_email_data(email) for email in result.get('value', [])]
                print(f"Successfully fetched {len(emails)} emails")
                return emails
            else:
                print(f"Error fetching emails: {response.status_code}")
                print(response.text)
                
                # If we got a complex filter error or any error, try simpler approaches
                if "InefficientFilter" in response.text or response.status_code != 200:
                    print("Trying fallback approach with local filtering")
                    # Try getting emails without filters and then filter locally
                    return self._get_emails_with_local_filtering(user_email, count, filters)
                    
                return []
                
        except Exception as e:
            print(f"Error retrieving emails: {str(e)}")
            return []
    
    def _format_email_data(self, email: Dict[str, Any]) -> Dict[str, Any]:
        """Format raw email data into a consistent structure"""
        return {
            "id": email.get("id", ""),
            "subject": email.get("subject", ""),
            "sender": email.get("sender", {}),
            "receivedDateTime": email.get("receivedDateTime", ""),
            "conversationId": email.get("conversationId", "")
        }
    
    def _search_by_sender(self, user_email: str, sender: str, count: int = 20, 
                        other_filters: Optional[Dict[str, str]] = None) -> List[Dict[str, Any]]:
        """Use $search for sender filtering which is more efficient"""
        if not self.access_token:
            self.access_token = self.get_access_token()
            
        if not self.access_token:
            return []
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        base_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        
        # Try different search syntaxes (Graph API can be picky about search syntax)
        search_attempts = [
            f'"from:{sender}"',  # Double quoted syntax
            f'from:{sender}',    # Unquoted syntax
            f'"{sender}"'        # Just the email in quotes
        ]
        
        emails = []
        for search_val in search_attempts:
            params = {
                "$select": "id,subject,sender,receivedDateTime,conversationId",
                "$top": count,
                "$search": search_val,
                "$orderby": "receivedDateTime desc"
            }
            
            # Add other filters if needed
            if other_filters:
                filter_parts = []
                
                # Subject filter (safe to add)
                if "subject" in other_filters and other_filters["subject"]:
                    filter_parts.append(f"contains(subject,'{other_filters['subject']}')")
                    
                # Date filters (generally safe to add)
                if "date_from" in other_filters and other_filters["date_from"]:
                    filter_parts.append(f"receivedDateTime ge {other_filters['date_from']}T00:00:00Z")
                    
                if "date_to" in other_filters and other_filters["date_to"]:
                    filter_parts.append(f"receivedDateTime le {other_filters['date_to']}T23:59:59Z")
                
                # Add filter if we have any parts (but be careful not to make it too complex)
                if filter_parts:
                    params["$filter"] = " and ".join(filter_parts)
            
            try:
                response = requests.get(base_url, headers=headers, params=params)
                
                if response.status_code == 200:
                    result = response.json()
                    found_emails = [self._format_email_data(email) for email in result.get('value', [])]
                    
                    if found_emails:
                        # If we found emails, return them
                        return found_emails
                        
                else:
                    print(f"Search attempt failed: {response.status_code}")
                    print(response.text)
                    
            except Exception as e:
                print(f"Error in search: {str(e)}")
        
        # If all search attempts fail, try fallback with local filtering
        return self._get_emails_with_local_filtering(user_email, count, {"from": sender, **other_filters} if other_filters else {"from": sender})
    
    def _get_emails_with_local_filtering(self, user_email: str, count: int, 
                                       filters: Optional[Dict[str, str]]) -> List[Dict[str, Any]]:
        """
        Fallback method to get emails without complex filters and filter locally
        This avoids the Graph API 'InefficientFilter' error
        """
        if not self.access_token:
            self.access_token = self.get_access_token()
            
        if not self.access_token:
            return []
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        base_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        
        # Use a simple query to get more emails than requested so we have room to filter
        params = {
            "$select": "id,subject,sender,receivedDateTime,conversationId,toRecipients",
            "$top": min(count * 3, 50),  # Get more emails but stay within limits
            "$orderby": "receivedDateTime desc"
        }
        
        # We can safely add date filters as they're usually efficient
        filter_parts = []
        if filters:
            if "date_from" in filters and filters["date_from"]:
                filter_parts.append(f"receivedDateTime ge {filters['date_from']}T00:00:00Z")
                
            if "date_to" in filters and filters["date_to"]:
                filter_parts.append(f"receivedDateTime le {filters['date_to']}T23:59:59Z")
        
        if filter_parts:
            params["$filter"] = " and ".join(filter_parts)
        
        try:
            response = requests.get(base_url, headers=headers, params=params)
            
            if response.status_code == 200:
                result = response.json()
                all_emails = result.get('value', [])
                filtered_emails = []
                
                # Apply local filtering
                for email in all_emails:
                    matches = True
                    
                    if filters:
                        # Check "from" filter
                        if "from" in filters and filters["from"]:
                            sender_address = email.get("sender", {}).get("emailAddress", {}).get("address", "").lower()
                            if not sender_address or filters["from"].lower() not in sender_address:
                                matches = False
                                
                        # Check "subject" filter
                        if "subject" in filters and filters["subject"]:
                            subject = email.get("subject", "").lower()
                            if not subject or filters["subject"].lower() not in subject:
                                matches = False
                                
                        # Check "to" filter
                        if "to" in filters and filters["to"]:
                            recipients = email.get("toRecipients", [])
                            recipient_match = False
                            for recipient in recipients:
                                address = recipient.get("emailAddress", {}).get("address", "").lower()
                                if filters["to"].lower() in address:
                                    recipient_match = True
                                    break
                            if not recipient_match:
                                matches = False
                    
                    if matches:
                        filtered_emails.append(self._format_email_data(email))
                        if len(filtered_emails) >= count:
                            break
                            
                return filtered_emails
                
            else:
                print(f"Error fetching emails: {response.status_code}")
                print(response.text)
                return []
                
        except Exception as e:
            print(f"Error retrieving emails: {str(e)}")
            return []
    
    def _get_clean_subject(self, subject: str) -> str:
        """
        Clean a subject line by removing prefixes like Re:, Fwd:, etc.
        
        Args:
            subject: The original subject line
            
        Returns:
            A cleaned subject line
        """
        if not subject:
            return "No Subject"
            
        # Define prefixes to remove (case insensitive)
        prefixes = [
            'RE:', 'FW:', 'FWD:', 'Re:', 'Fw:', 'Fwd:',
            'RE: ', 'FW: ', 'FWD: ', 'Re: ', 'Fw: ', 'Fwd: '
        ]
        
        # Remove prefixes iteratively until no more are found
        cleaned = subject
        changed = True
        while changed:
            prev_length = len(cleaned)
            for prefix in prefixes:
                if cleaned.startswith(prefix):
                    cleaned = cleaned[len(prefix):].strip()
            changed = len(cleaned) < prev_length
            
        return cleaned if cleaned else "No Subject"
    
    def get_email_thread(self, user_email, conversation_id):
        """
        Get all emails in a thread based on conversationId
        
        Args:
            user_email (str): The email address of the user
            conversation_id (str): The conversation ID to fetch
            
        Returns:
            dict: A dictionary with two keys:
                - 'messages': A list of email objects sorted chronologically
                - 'thread_structure': A nested tree structure of emails built using 
                                      internetMessageId and inReplyTo fields
        """
        if not self.access_token:
            self.access_token = self.get_access_token()
            
        if not self.access_token:
            print("No access token available. Authentication failed.")
            return {'messages': [], 'thread_structure': []}
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "Prefer": "outlook.allow-unsafe-html,outlook.body-content-type=html"
        }
        
        # First get all message IDs in the conversation
        base_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        
        try:
            # Try multiple methods to ensure we get all messages in the thread
            complete_messages = []
            
            # Method 1: Direct filter by conversationId with all important fields
            print(f"\n=== Method 1: Direct filter by conversationId ===")
            params = {
                "$select": "id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender,internetMessageId,inReplyTo",
                "$filter": f"conversationId eq '{conversation_id}'",
                "$top": 100,  # Increase to get more emails in the thread
                "$orderby": "receivedDateTime asc"  # Get in chronological order
            }
            
            print(f"Fetching thread with ID: {conversation_id}")
            print(f"Request: {base_url}")
            print(f"Params: {params}")
            
            response = requests.get(base_url, headers=headers, params=params)
            
            print(f"Response status: {response.status_code}")
            
            if response.status_code == 200:
                thread_messages = response.json().get('value', [])
                print(f"Method 1 found {len(thread_messages)} messages in thread")
                
                if thread_messages:
                    # First, check if we need to fetch full content for any of the messages
                    for msg in thread_messages:
                        if not msg.get('body', {}).get('content'):
                            # Need to fetch full content
                            msg_id = msg.get('id')
                            print(f"Need to fetch full content for message: {msg.get('subject', 'No Subject')}")
                            try:
                                complete_msg = self._get_email_with_body(user_email, msg_id)
                                if complete_msg:
                                    complete_messages.append(complete_msg)
                                    print(f"Added message with full content: {complete_msg.get('subject', 'No Subject')}")
                            except Exception as e:
                                print(f"Error fetching complete message {msg_id}: {str(e)}")
                                # Still add the original message without full body
                                complete_messages.append(msg)
                        else:
                            # Already has content, just add it
                            complete_messages.append(msg)
                            print(f"Added message (already has content): {msg.get('subject', 'No Subject')}")
            
            # If first method failed or found few messages, try method 2
            if len(complete_messages) < 2:
                print(f"\n=== Method 2: Use API's dedicated thread endpoint ===")
                
                # Try Microsoft Graph beta endpoint for retrieving conversations
                # This endpoint might be more reliable for thread retrieval
                try:
                    # New approach: first get all conversation participants
                    beta_participants_url = f"https://graph.microsoft.com/beta/users/{user_email}/messages/{complete_messages[0]['id'] if complete_messages else conversation_id}/conversationParticipants"
                    print(f"Trying to get conversation participants: {beta_participants_url}")
                    
                    participants_response = requests.get(beta_participants_url, headers=headers)
                    participants = []
                    
                    if participants_response.status_code == 200:
                        participants_data = participants_response.json()
                        participants = [p.get('emailAddress', {}).get('address', '') 
                                      for p in participants_data.get('value', [])]
                        print(f"Found participants: {participants}")
                    
                    # Then query with standard and beta endpoints
                    conversation_url = f"https://graph.microsoft.com/beta/users/{user_email}/messages?$filter=conversationId eq '{conversation_id}'&$top=100&$select=id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender&$orderby=receivedDateTime asc"
                    
                    print(f"Trying beta endpoint: {conversation_url}")
                    beta_response = requests.get(conversation_url, headers=headers)
                    
                    if beta_response.status_code == 200:
                        beta_messages = beta_response.json().get('value', [])
                        print(f"Beta endpoint found {len(beta_messages)} messages")
                        
                        # Remove duplicates by ID
                        existing_ids = set(msg.get('id') for msg in complete_messages)
                        for msg in beta_messages:
                            if msg.get('id') not in existing_ids:
                                # Fetch full content if needed
                                if not msg.get('body', {}).get('content'):
                                    try:
                                        complete_msg = self._get_email_with_body(user_email, msg.get('id'))
                                        if complete_msg:
                                            complete_messages.append(complete_msg)
                                            existing_ids.add(msg.get('id'))
                                            print(f"Added message from beta API: {complete_msg.get('subject', 'No Subject')}")
                                    except Exception as e:
                                        print(f"Error fetching complete message: {str(e)}")
                                        complete_messages.append(msg)
                                        existing_ids.add(msg.get('id'))
                                else:
                                    complete_messages.append(msg)
                                    existing_ids.add(msg.get('id'))
                                    print(f"Added message from beta API: {msg.get('subject', 'No Subject')}")
                    
                    # Also try a search across all participants if we have them
                    if participants:
                        print(f"Searching for messages involving all participants")
                        for participant in participants:
                            if participant and participant != user_email:
                                participant_filter = f"from/emailAddress/address eq '{participant}' or recipients/emailAddress/address has '{participant}'"
                                participant_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages?$filter={participant_filter}&$top=50&$select=id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender"
                                
                                try:
                                    participant_response = requests.get(participant_url, headers=headers)
                                    if participant_response.status_code == 200:
                                        participant_messages = participant_response.json().get('value', [])
                                        print(f"Found {len(participant_messages)} messages involving {participant}")
                                        
                                        existing_ids = set(msg.get('id') for msg in complete_messages)
                                        for msg in participant_messages:
                                            if msg.get('id') not in existing_ids:
                                                # Check for subject similarity to include in thread
                                                original_subject = complete_messages[0].get('subject', '') if complete_messages else ''
                                                clean_original = self._get_clean_subject(original_subject)
                                                clean_msg = self._get_clean_subject(msg.get('subject', ''))
                                                
                                                # Add if subjects match or has same conversation ID
                                                if msg.get('conversationId') == conversation_id or (clean_original and clean_msg and clean_original.lower() == clean_msg.lower()):
                                                    try:
                                                        complete_msg = self._get_email_with_body(user_email, msg.get('id'))
                                                        if complete_msg:
                                                            complete_messages.append(complete_msg)
                                                            existing_ids.add(msg.get('id'))
                                                            print(f"Added participant message: {complete_msg.get('subject', 'No Subject')}")
                                                    except Exception as e:
                                                        print(f"Error fetching participant message: {str(e)}")
                                except Exception as e:
                                    print(f"Error searching participant messages: {str(e)}")
                except Exception as e:
                    print(f"Error using beta endpoint: {str(e)}")
            
            # Method 3: Search by similar subjects if we still don't have at least 2 messages
            if len(complete_messages) < 2:
                print(f"\n=== Method 3: Try to find related messages by subject ===")
                # Get a sample message to extract subject
                sample_msg = None
                if complete_messages:
                    sample_msg = complete_messages[0]
                
                if sample_msg and 'subject' in sample_msg:
                    # Get subject without prefixes
                    subject = sample_msg.get('subject', '')
                    clean_subject = self._get_clean_subject(subject)
                    
                    if clean_subject:
                        # Search by cleaned subject
                        print(f"Searching for related messages with subject: {clean_subject}")
                        subject_params = {
                            "$select": "id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender",
                            "$search": f"\"{clean_subject}\"", 
                            "$top": 50,
                            "$orderby": "receivedDateTime asc"
                        }
                        
                        response = requests.get(base_url, headers=headers, params=subject_params)
                        
                        if response.status_code == 200:
                            subject_messages = response.json().get('value', [])
                            print(f"Method 3 found {len(subject_messages)} messages with similar subject")
                            
                            # Remove duplicates by ID
                            existing_ids = set(msg.get('id') for msg in complete_messages)
                            for msg in subject_messages:
                                if msg.get('id') not in existing_ids:
                                    # Only add if it might be related (has same conversation ID or very similar subject)
                                    msg_subject = self._get_clean_subject(msg.get('subject', ''))
                                    if msg.get('conversationId') == conversation_id or clean_subject.lower() in msg_subject.lower():
                                        # Get full content if needed
                                        if not msg.get('body', {}).get('content'):
                                            try:
                                                complete_msg = self._get_email_with_body(user_email, msg.get('id'))
                                                if complete_msg:
                                                    complete_messages.append(complete_msg)
                                                    existing_ids.add(msg.get('id'))
                                                    print(f"Added related message: {complete_msg.get('subject', 'No Subject')}")
                                            except Exception as e:
                                                print(f"Error fetching complete message: {str(e)}")
                                                complete_messages.append(msg)
                                                existing_ids.add(msg.get('id'))
                                        else:
                                            complete_messages.append(msg)
                                            existing_ids.add(msg.get('id'))
                                            print(f"Added related message: {msg.get('subject', 'No Subject')}")
            
            # Method 4: Check if we specifically need to look for "Driving License" thread
            if any("driving license" in msg.get('subject', '').lower() for msg in complete_messages) and len(complete_messages) < 2:
                print(f"\n=== Method 4: Special case for Driving License thread ===")
                driving_license_params = {
                    "$select": "id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender",
                    "$search": "\"Driving License\"", 
                    "$top": 10
                }
                
                response = requests.get(base_url, headers=headers, params=driving_license_params)
                if response.status_code == 200:
                    dl_messages = response.json().get('value', [])
                    print(f"Found {len(dl_messages)} Driving License messages")
                    
                    # Remove duplicates by ID
                    existing_ids = set(msg.get('id') for msg in complete_messages)
                    for msg in dl_messages:
                        if msg.get('id') not in existing_ids:
                            try:
                                complete_msg = self._get_email_with_body(user_email, msg.get('id'))
                                if complete_msg:
                                    complete_messages.append(complete_msg)
                                    existing_ids.add(msg.get('id'))
                                    print(f"Added Driving License message: {complete_msg.get('subject', 'No Subject')}")
                            except Exception as e:
                                print(f"Error fetching Driving License message: {str(e)}")
            
            # Sort all messages by date - emails should be in chronological order
            complete_messages.sort(key=lambda m: m.get('receivedDateTime', ''))
            
            print(f"\n=== Thread Results Summary ===")
            print(f"Retrieved {len(complete_messages)} total messages from thread")
            for idx, msg in enumerate(complete_messages):
                print(f"  {idx+1}. {msg.get('subject', 'No Subject')} - {msg.get('receivedDateTime', 'Unknown')}")
            
            # Build the structured thread representation
            thread_structure = self.build_thread_structure(complete_messages)
            
            # Inject the thread structure as an additional property
            return {
                'messages': complete_messages,  # Original chronological list for backward compatibility
                'thread_structure': thread_structure  # New structured representation with parent-child relationships
            }
            
        except Exception as e:
            print(f"Error retrieving thread: {str(e)}")
            import traceback
            traceback.print_exc()
            return {'messages': [], 'thread_structure': []}
    
    def simulate_thread_emails(self, original_email):
        """
        Create simulated thread emails based on an original email for testing
        
        Args:
            original_email: Original email dict from the Graph API
            
        Returns:
            dict: A dictionary with the same format as get_email_thread:
                - 'messages': A list of email objects sorted chronologically
                - 'thread_structure': A structured thread representation
        """
        if not original_email:
            return {'messages': [], 'thread_structure': []}
            
        # Clone the original email to avoid modifying it
        result = [original_email]
        
        # Ensure original email has internetMessageId if not present
        if 'internetMessageId' not in original_email:
            original_email['internetMessageId'] = f"original_{original_email.get('id', 'unknown')}@example.com"
        
        # Add sender details for creating the reply
        sender_address = original_email.get('sender', {}).get('emailAddress', {})
        sender_name = sender_address.get('name', 'Unknown')
        sender_email = sender_address.get('address', 'unknown@example.com')
        
        # First reply - let's say it's sent 1 day later
        date_str = original_email.get('receivedDateTime', '')
        try:
            # Try to parse the date
            from datetime import datetime, timedelta
            date_obj = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            reply_date = (date_obj + timedelta(days=1)).isoformat().replace('+00:00', 'Z')
        except Exception:
            # Fallback if date parsing fails
            reply_date = original_email.get('receivedDateTime', '')
            
        # Create a reply from the original sender
        subject = original_email.get('subject', '')
        if not subject.startswith('Re:'):
            subject = 'Re: ' + subject
            
        # Generate a reply body based on the original email content
        original_body = original_email.get('body', {}).get('content', '')
        reply_body = f"""
        <html>
        <body>
        <p>Thank you for your interest. I can confirm we have received your request.</p>
        <p>We'll process this and get back to you with more details soon.</p>
        <p>Best regards,<br>
        {sender_name}</p>
        <hr>
        <div style="color:#666">
        Original message:<br>
        {original_body}
        </div>
        </body>
        </html>
        """
        
        # Create the reply email object with threading information
        reply = {
            'id': original_email.get('id', '') + '_reply1',
            'subject': subject,
            'sender': original_email.get('sender', {}),
            'receivedDateTime': reply_date,
            'conversationId': original_email.get('conversationId', ''),
            'body': {
                'contentType': 'html',
                'content': reply_body
            },
            'internetMessageId': f"reply_{original_email.get('id', 'unknown')}@example.com",
            'inReplyTo': original_email.get('internetMessageId')
        }
        
        result.append(reply)
        
        # Build thread structure
        thread_structure = self.build_thread_structure(result)
        
        return {
            'messages': result,
            'thread_structure': thread_structure
        }

    def _get_email_with_body(self, user_email: str, email_id: str) -> Dict[str, Any]:
        """Get a single email with full body content"""
        if not self.access_token:
            self.access_token = self.get_access_token()
            
        if not self.access_token:
            return {}
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "Prefer": "outlook.allow-unsafe-html,outlook.body-content-type=html"
        }
        
        email_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages/{email_id}"
        params = {
            "$select": "id,subject,sender,from,receivedDateTime,body,bodyPreview,toRecipients,ccRecipients,conversationId,internetMessageId,inReplyTo"
        }
        
        try:
            print(f"Fetching email details for ID: {email_id}")
            response = requests.get(email_url, headers=headers, params=params)
            
            if response.status_code == 200:
                email_data = response.json()
                
                # Ensure both sender and from fields are present
                if 'from' in email_data and 'sender' not in email_data:
                    email_data['sender'] = email_data['from']
                elif 'sender' in email_data and 'from' not in email_data:
                    email_data['from'] = email_data['sender']
                    
                # Verify we have the body content    
                if 'body' not in email_data or not email_data.get('body', {}).get('content'):
                    print(f"Warning: Email {email_id} has empty or missing body content")
                
                print(f"Successfully fetched email with subject: {email_data.get('subject', 'No Subject')}")
                return email_data
            else:
                print(f"Error fetching email detail: {response.status_code}")
                print(response.text)
                return {}
        except Exception as e:
            print(f"Error retrieving email detail: {str(e)}")
            return {}

    def build_thread_structure(self, messages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Build a structured thread representation using internetMessageId and inReplyTo fields.
        
        Args:
            messages: List of email messages from the Graph API with internetMessageId and inReplyTo fields
            
        Returns:
            List of thread root messages, with nested 'replies' fields containing child messages
        """
        if not messages:
            return []
            
        # First, build a dictionary of messages by internetMessageId for quick lookup
        by_id = {}
        for msg in messages:
            internet_msg_id = msg.get('internetMessageId')
            if internet_msg_id:
                by_id[internet_msg_id] = msg
                # Initialize replies field for each message
                msg['replies'] = []
                
        # Build the thread structure by connecting messages via inReplyTo references
        thread_roots = []
        
        for msg in messages:
            # Get the parent message ID (inReplyTo)
            parent_id = msg.get('inReplyTo')
            internet_msg_id = msg.get('internetMessageId')
            
            # Skip if this message doesn't have an internetMessageId (unlikely but possible)
            if not internet_msg_id:
                continue
                
            if parent_id and parent_id in by_id:
                # This message is a reply to another message in our set
                parent = by_id[parent_id]
                parent['replies'].append(msg)
            else:
                # This is a root message (no parent or parent not in our set)
                # Only add to roots if not already a child of another message
                # Check if this message is already a child of another message
                is_child = False
                for potential_parent in by_id.values():
                    if msg in potential_parent.get('replies', []):
                        is_child = True
                        break
                        
                if not is_child:
                    thread_roots.append(msg)
        
        # If we end up with no thread roots (rare edge case, but possible),
        # fall back to using the original chronological order
        if not thread_roots and messages:
            # Sort by date
            messages.sort(key=lambda m: m.get('receivedDateTime', ''))
            # Return the earliest message as the root
            thread_roots = [messages[0]]
            
        return thread_roots
        
    def get_messages_with_threading_info(self, user_email: str, count: int = 100):
        """Get messages with internetMessageId and inReplyTo fields to reconstruct threads."""
        if not self.access_token:
            self.access_token = self.get_access_token()
            
        if not self.access_token:
            print("No access token available. Authentication failed.")
            return []
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "Prefer": "outlook.allow-unsafe-html,outlook.body-content-type=html"
        }
        
        # Request specific fields needed for thread reconstruction
        base_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        params = {
            "$select": "id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender,internetMessageId,inReplyTo",
            "$top": count,
            "$orderby": "receivedDateTime desc"
        }
        
        try:
            print(f"Fetching messages with threading info...")
            response = requests.get(base_url, headers=headers, params=params)
            
            if response.status_code == 200:
                messages = response.json().get('value', [])
                print(f"Retrieved {len(messages)} messages with threading info")
                return messages
            else:
                print(f"Error fetching messages: {response.status_code}")
                print(response.text)
                return []
        except Exception as e:
            print(f"Error retrieving messages: {str(e)}")
            return []
            
    def get_current_user_email(self):
        """Get email address of current user."""
        if not self.access_token:
            self.access_token = self.get_access_token()
            
        if not self.access_token:
            return None
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
            
            if response.status_code == 200:
                user_data = response.json()
                return user_data.get('userPrincipalName') or user_data.get('mail')
            else:
                print(f"Error getting current user: {response.status_code}")
                return None
        except Exception as e:
            print(f"Error retrieving current user: {str(e)}")
            return None
            
    def get_thread_messages(self, message_id):
        """Get all messages in a thread based on a message ID."""
        print(f"Getting all messages in thread containing message: {message_id}")
        
        # First get the specified message
        user_email = self.get_current_user_email()
        target_message = self._get_email_with_body(user_email, message_id)
        
        if not target_message:
            print("Could not retrieve target message")
            return []
        
        # Get threading information
        messages = self.get_messages_with_threading_info(user_email)
        
        # Build thread structure
        thread_roots = self.build_thread_structure(messages)
        
        # Find which thread contains our target message
        target_thread = None
        for root in thread_roots:
            if self._message_in_thread(root, message_id):
                target_thread = root
                break
        
        if not target_thread:
            print(f"Could not find thread containing message: {message_id}")
            # Fallback to conversation id if available
            conversation_id = target_message.get('conversationId')
            if conversation_id:
                print(f"Falling back to conversationId: {conversation_id}")
                thread_result = self.get_email_thread(user_email, conversation_id)
                return thread_result.get('messages', [])
            return [target_message]
        
        # Collect all messages in the thread
        thread_messages = []
        self._collect_thread_messages(target_thread, thread_messages)
        
        # Sort by date
        thread_messages.sort(key=lambda m: m.get('receivedDateTime', ''))
        
        print(f"Retrieved {len(thread_messages)} messages in thread")
        return thread_messages
        
    def _message_in_thread(self, thread_root, message_id):
        """Recursively check if message_id is in the thread rooted at thread_root."""
        if thread_root.get('id') == message_id:
            return True
        
        for reply in thread_root.get('replies', []):
            if self._message_in_thread(reply, message_id):
                return True
        
        return False
        
    def _collect_thread_messages(self, message, result_list):
        """Recursively collect all messages in a thread starting at message."""
        result_list.append(message)
        
        for reply in message.get('replies', []):
            self._collect_thread_messages(reply, result_list)
    
    def group_emails_by_threads(self, user_email: str, count: int = 100) -> List[Dict[str, Any]]:
        """
        Group emails into threads using internetMessageId and inReplyTo fields
        
        This approach is more accurate than using conversationId because it can
        find related emails even when their conversation IDs are different.
        
        Args:
            user_email: User email address
            count: Maximum number of emails to fetch
            
        Returns:
            List of thread objects containing thread information and message IDs
        """
        if not self.access_token:
            self.access_token = self.get_access_token()
            if not self.access_token:
                return []
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "Prefer": "outlook.allow-unsafe-html,outlook.body-content-type=html"
        }
        
        # First get recent emails to analyze
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        params = {
            "$select": "id,subject,conversationId,receivedDateTime,from,toRecipients,internetMessageId,inReplyTo,sender",
            "$top": count,
            "$orderby": "receivedDateTime desc"
        }
        
        try:
            print(f"\n=== Fetching {count} recent emails for thread analysis ===")
            response = requests.get(url, headers=headers, params=params)
            
            if response.status_code == 200:
                data = response.json()
                messages = data.get('value', [])
                print(f"Retrieved {len(messages)} messages")
                
                # Initialize structures for thread grouping
                messages_by_id = {}  # Map internetMessageId -> message
                threads = {}         # Map thread_id -> thread info
                
                # First pass: build a map of messages by internetMessageId
                for msg in messages:
                    internet_msg_id = msg.get('internetMessageId')
                    if internet_msg_id:
                        messages_by_id[internet_msg_id] = msg
                
                # Second pass: build thread relationships
                for msg in messages:
                    internet_msg_id = msg.get('internetMessageId')
                    in_reply_to = msg.get('inReplyTo')
                    subject = msg.get('subject', 'No Subject')
                    clean_subject = self._get_clean_subject(subject)
                    
                    # Skip messages without internetMessageId
                    if not internet_msg_id:
                        continue
                    
                    # Check if this message is already in a thread
                    thread_id = None
                    for tid, thread in threads.items():
                        if internet_msg_id in thread['message_ids'] or (in_reply_to and in_reply_to in thread['message_ids']):
                            thread_id = tid
                            break
                    
                    # If not in a thread, create a new one
                    if not thread_id:
                        thread_id = f"thread_{len(threads) + 1}"
                        threads[thread_id] = {
                            'thread_id': thread_id,
                            'subject': clean_subject,
                            'message_count': 0,
                            'message_ids': set(),
                            'messages': [],
                            'first_message_date': msg.get('receivedDateTime', ''),
                            'latest_message_date': msg.get('receivedDateTime', ''),
                            'participants': set()
                        }
                    
                    # Add this message to the thread
                    threads[thread_id]['message_ids'].add(internet_msg_id)
                    threads[thread_id]['messages'].append(msg.get('id'))
                    threads[thread_id]['message_count'] += 1
                    
                    # Update thread dates
                    msg_date = msg.get('receivedDateTime', '')
                    if msg_date < threads[thread_id]['first_message_date']:
                        threads[thread_id]['first_message_date'] = msg_date
                    if msg_date > threads[thread_id]['latest_message_date']:
                        threads[thread_id]['latest_message_date'] = msg_date
                    
                    # Add participants
                    if 'from' in msg and 'emailAddress' in msg['from'] and 'address' in msg['from']['emailAddress']:
                        threads[thread_id]['participants'].add(msg['from']['emailAddress']['address'])
                    
                    # If this message is a reply, add the parent to the thread too
                    if in_reply_to and in_reply_to in messages_by_id:
                        parent_msg = messages_by_id[in_reply_to]
                        threads[thread_id]['message_ids'].add(in_reply_to)
                        threads[thread_id]['messages'].append(parent_msg.get('id'))
                        threads[thread_id]['message_count'] += 1
                        
                        # Update first message date if parent is earlier
                        parent_date = parent_msg.get('receivedDateTime', '')
                        if parent_date < threads[thread_id]['first_message_date']:
                            threads[thread_id]['first_message_date'] = parent_date
                
                # Convert to list and finalize
                thread_list = list(threads.values())
                for thread in thread_list:
                    # Remove duplicates from messages list
                    thread['messages'] = list(set(thread['messages']))
                    thread['message_count'] = len(thread['messages'])
                    # Convert participants set to list
                    thread['participants'] = list(thread['participants'])
                
                # Sort threads by latest message date
                thread_list.sort(key=lambda t: t['latest_message_date'], reverse=True)
                
                print(f"Grouped emails into {len(thread_list)} threads using internetMessageId and inReplyTo fields")
                return thread_list
            else:
                print(f"Error fetching messages: {response.status_code}")
                print(response.text)
                return []
        except Exception as e:
            print(f"Exception in group_emails_by_threads: {str(e)}")
            return []

    def fetch_thread_messages(self, user_email: str, thread_id: str) -> Dict[str, Any]:
        """
        Fetch all messages for a thread based on the thread ID
        
        This method handles threads with mixed conversation IDs by using internetMessageId
        and inReplyTo fields to ensure all related messages are included.
        
        Args:
            user_email: User email address
            thread_id: The thread ID to fetch
            
        Returns:
            dict: A dictionary with keys:
                - 'messages': List of email messages
                - 'thread_structure': Hierarchical structure of the thread
        """
        import re
        
        if not self.access_token:
            self.access_token = self.get_access_token()
            if not self.access_token:
                return {'messages': [], 'thread_structure': []}
                
        # First try to get messages using the thread ID as a conversation ID
        thread_result = self.get_email_thread(user_email, thread_id)
        
        # Check if we got a valid result
        if not thread_result or not isinstance(thread_result, dict) or not thread_result.get('messages'):
            # If not, try to see if this is a message ID instead of a conversation ID
            print(f"Thread ID {thread_id} didn't work as a conversation ID, trying as a message ID")
            message = self._get_email_with_body(user_email, thread_id)
            
            if message:
                # Get related messages by looking for the same internetMessageId pattern
                # or by checking references/inReplyTo fields
                internet_message_id = message.get('internetMessageId')
                
                if internet_message_id:
                    # Find all messages with the same conversation topic
                    related_messages = self.find_related_messages(user_email, internet_message_id)
                    
                    if related_messages:
                        # Clean the message bodies using our improved approach
                        for msg in related_messages:
                            if 'body' in msg and 'content' in msg['body']:
                                content = msg['body']['content']
                                # Find "from:" in case insensitive way
                                from_match = re.search(r'[Ff][Rr][Oo][Mm]:', content)
                                if from_match:
                                    # Keep only content before "from:"
                                    content = content[:from_match.start()]
                                    msg['body']['content'] = content
                        
                        # Build the thread structure
                        thread_structure = self.build_thread_structure(related_messages)
                        return {
                            'messages': related_messages,
                            'thread_structure': thread_structure
                        }
                    else:
                        # Just return the single message we found
                        return {
                            'messages': [message],
                            'thread_structure': [message]
                        }
            
            # If all else fails, return empty result
            return {'messages': [], 'thread_structure': []}
        
        # Clean message bodies in the thread result
        if thread_result and 'messages' in thread_result:
            for msg in thread_result['messages']:
                if 'body' in msg and 'content' in msg['body']:
                    content = msg['body']['content']
                    # Find "from:" in case insensitive way
                    from_match = re.search(r'[Ff][Rr][Oo][Mm]:', content)
                    if from_match:
                        # Keep only content before "from:"
                        content = content[:from_match.start()]
                        msg['body']['content'] = content
        
        return thread_result

    def find_related_messages(self, user_email: str, internet_message_id: str) -> List[Dict[str, Any]]:
        """
        Find all messages related to a given internetMessageId
        
        This searches for:
        1. Messages with the same internetMessageId
        2. Messages that have this internetMessageId in their inReplyTo field
        3. Messages that this message refers to in its inReplyTo field
        
        Args:
            user_email: User email address
            internet_message_id: The internetMessageId to search for
            
        Returns:
            List of related email messages
        """
        if not self.access_token:
            self.access_token = self.get_access_token()
            if not self.access_token:
                return []
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "Prefer": "outlook.allow-unsafe-html,outlook.body-content-type=html"
        }
        
        # Remove any angle brackets if present in the internetMessageId
        clean_id = internet_message_id.strip('<>')
        
        # Build the URL to search for related messages
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        
        # Search for messages with the same internetMessageId or 
        # messages that reference this ID in their inReplyTo field
        params = {
            "$select": "id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender,internetMessageId,inReplyTo",
            "$filter": f"internetMessageId eq '{clean_id}' or contains(inReplyTo, '{clean_id}')",
            "$top": 100,
            "$orderby": "receivedDateTime asc"
        }
        
        try:
            response = requests.get(url, headers=headers, params=params)
            
            if response.status_code == 200:
                data = response.json()
                messages = data.get('value', [])
                
                # Also find the message this one is replying to, if any
                reply_to_messages = []
                for msg in messages:
                    in_reply_to = msg.get('inReplyTo')
                    if in_reply_to:
                        # Clean up the inReplyTo value
                        in_reply_to = in_reply_to.strip('<>')
                        
                        # Find the message this is replying to
                        parent_params = {
                            "$select": "id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender,internetMessageId,inReplyTo",
                            "$filter": f"internetMessageId eq '{in_reply_to}'",
                            "$top": 1
                        }
                        
                        parent_response = requests.get(url, headers=headers, params=parent_params)
                        
                        if parent_response.status_code == 200:
                            parent_data = parent_response.json()
                            parent_messages = parent_data.get('value', [])
                            
                            # Add any found parent messages that aren't already in our list
                            for parent in parent_messages:
                                parent_id = parent.get('id')
                                if parent_id and not any(m.get('id') == parent_id for m in messages):
                                    reply_to_messages.append(parent)
                
                # Combine all found messages
                all_messages = messages + reply_to_messages
                
                # Ensure we have full body content for each message
                for i, msg in enumerate(all_messages):
                    if not msg.get('body', {}).get('content'):
                        try:
                            full_msg = self._get_email_with_body(user_email, msg.get('id'))
                            if full_msg:
                                all_messages[i] = full_msg
                        except Exception as e:
                            print(f"Error fetching full message: {str(e)}")
                
                # Sort by date
                all_messages.sort(key=lambda m: m.get('receivedDateTime', ''))
                return all_messages
            else:
                print(f"Error searching for related messages: {response.status_code}")
                print(response.text)
                return []
        except Exception as e:
            print(f"Exception searching for related messages: {str(e)}")
            return []

    def group_emails_by_subject(self, user_email: str, count: int = 100) -> List[Dict[str, Any]]:
        """
        Group emails into threads based on subject similarity, treating Re: emails as replies.
        This approach groups emails even when conversationId is different.
        
        Args:
            user_email: User email address
            count: Maximum number of emails to fetch
            
        Returns:
            List of thread objects containing thread information and message IDs
        """
        if not self.access_token:
            self.access_token = self.get_access_token()
            if not self.access_token:
                return []
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "Prefer": "outlook.allow-unsafe-html,outlook.body-content-type=html"
        }
        
        # First get recent emails to analyze
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        params = {
            "$select": "id,subject,conversationId,receivedDateTime,from,toRecipients,body,sender",
            "$top": count,
            "$orderby": "receivedDateTime desc"
        }
        
        try:
            print(f"\n=== Fetching {count} recent emails for subject-based thread analysis ===")
            response = requests.get(url, headers=headers, params=params)
            
            if response.status_code == 200:
                data = response.json()
                messages = data.get('value', [])
                print(f"Retrieved {len(messages)} messages")
                
                # Clean up emails by extracting body content
                for msg in messages:
                    if 'body' in msg and 'content' in msg['body']:
                        # Clean up metadata by removing content after "from:"
                        content = msg['body']['content']
                        # Find "from:" in case insensitive way
                        from_match = re.search(r'[Ff][Rr][Oo][Mm]:', content)
                        if from_match:
                            # Keep only content before "from:"
                            content = content[:from_match.start()]
                        msg['cleaned_body'] = content
                
                # Dictionary to hold subject-based threads
                threads = {}
                
                # Process messages and group by clean subject
                for msg in messages:
                    subject = msg.get('subject', 'No Subject')
                    # Clean the subject (remove Re:, Fw:, etc.)
                    clean_subject = self._get_clean_subject(subject)
                    
                    # Find a matching thread by subject similarity
                    match_found = False
                    for thread_key in list(threads.keys()):
                        # Use fuzzy matching to detect similar subjects (85% similarity threshold)
                        similarity = fuzz.ratio(thread_key.lower(), clean_subject.lower())
                        if similarity >= 85:
                            # Add this message to existing thread
                            threads[thread_key]['messages'].append(msg.get('id'))
                            threads[thread_key]['message_count'] += 1
                            # Update thread dates if necessary
                            msg_date = msg.get('receivedDateTime', '')
                            if msg_date > threads[thread_key]['latest_message_date']:
                                threads[thread_key]['latest_message_date'] = msg_date
                                threads[thread_key]['latest_message'] = msg
                            if msg_date < threads[thread_key]['first_message_date']:
                                threads[thread_key]['first_message_date'] = msg_date
                                threads[thread_key]['first_message'] = msg
                            
                            # Add participants if not already included
                            sender_address = msg.get('from', {}).get('emailAddress', {})
                            sender_email = sender_address.get('address', '')
                            if sender_email and sender_email not in threads[thread_key]['participants']:
                                threads[thread_key]['participants'].append(sender_email)
                            
                            match_found = True
                            break
                    
                    # If no match found, create a new thread
                    if not match_found:
                        thread_id = f"thread_{len(threads) + 1}"
                        threads[clean_subject] = {
                            'thread_id': thread_id,
                            'subject': clean_subject,
                            'original_subject': subject,
                            'message_count': 1,
                            'messages': [msg.get('id')],
                            'first_message_date': msg.get('receivedDateTime', ''),
                            'latest_message_date': msg.get('receivedDateTime', ''),
                            'first_message': msg,
                            'latest_message': msg,
                            'participants': [msg.get('from', {}).get('emailAddress', {}).get('address', '')]
                        }
                
                # Convert threads dictionary to list
                thread_list = list(threads.values())
                
                # Sort threads by latest message date (newest first)
                thread_list.sort(key=lambda t: t['latest_message_date'], reverse=True)
                
                print(f"Grouped {len(messages)} emails into {len(thread_list)} threads based on subject similarity")
                return thread_list
            else:
                print(f"Error fetching messages: {response.status_code}")
                print(response.text)
                return []
        except Exception as e:
            print(f"Exception in group_emails_by_subject: {str(e)}")
            return []