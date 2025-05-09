"""
AI service for processing emails and generating groundtruth data.
"""

import json
import google.generativeai as genai
from typing import Dict, Any, List, Optional
from datetime import datetime
import requests
import logging
import re

class AIService:
    """
    AI service for analyzing emails using the Dwellworks API endpoint
    """
    
    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize the AI Service
        
        Args:
            api_key: Optional API key (not needed for Dwellworks endpoint)
        """
        self.api_endpoint = "https://mlemailintegrationservices-gef9fwepguapgwfr.eastus2-01.azurewebsites.net/test_extraction"
        # Configure logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger("AIService")
    
    def analyze_email_thread(self, emails: List[Dict[str, Any]], **kwargs) -> Dict[str, Any]:
        """
        Analyze an email using the Dwellworks API
        
        Args:
            emails: List of email dictionaries from Graph API
            **kwargs: Additional arguments including previous_summary
            
        Returns:
            API response from Dwellworks
        """
        if not emails:
            self.logger.error("No emails provided for analysis")
            raise ValueError("No emails provided for analysis")
        
        # Extract previous_summary from kwargs if available
        previous_summary = kwargs.get('previous_summary', '')
        
        # Format the email for the API
        email = emails[0]  # Only use the first email provided
        self.logger.info(f"Processing email with ID: {email.get('id', 'unknown')[:10]}...")
        
        # Process the email body to clean HTML
        body_content = email.get("body", {}).get("content", "")
        # Strip HTML tags if present
        cleaned_body = re.sub(r'<[^>]+>', ' ', body_content)
        # Remove excessive whitespace
        cleaned_body = re.sub(r'\s+', ' ', cleaned_body).strip()
        
        # Get sender details - safely extract from Graph API response structure
        sender_name = email.get("sender", {}).get("emailAddress", {}).get("name", "Unknown")
        sender_email = email.get("sender", {}).get("emailAddress", {}).get("address", "")
        
        # Create formatted data for API
        formatted_data = {
            "mail_id": email.get("id", ""),
            "file_name": [],
            "email": sender_email,
            "mail_time": email.get("receivedDateTime", ""),
            "body_type": "html",
            "mail_body": cleaned_body,
            "thread_id": email.get("conversationId", ""),
            "mail_summary": previous_summary
        }
        
        # Log the formatted data
        self.logger.info(f"Sending request to {self.api_endpoint}")
        self.logger.info(f"Request data: mail_id={formatted_data['mail_id'][:10]}..., from={formatted_data['email']}")
        self.logger.info(f"Request data contains {len(formatted_data['mail_body'])} characters of email body")
        if previous_summary:
            self.logger.info(f"Including previous summary: {previous_summary[:100]}...")
        
        try:
            # Send request to Dwellworks API
            response = requests.post(
                self.api_endpoint,
                json=formatted_data,
                headers={"Content-Type": "application/json"},
                timeout=30
            )
            
            # Check response status
            if response.status_code != 200:
                self.logger.error(f"API returned error status: {response.status_code}")
                self.logger.error(f"Response content: {response.text[:200]}...")
                response.raise_for_status()
            
            # Parse the response as JSON
            try:
                result = response.json()
                self.logger.info(f"Received API response with {len(str(result))} characters")
                
                # DO NOT post-process the result - keep the original values from the API
                # We'll handle normalization in the evaluation step instead
                return result
            except json.JSONDecodeError:
                self.logger.error(f"Failed to parse API response as JSON: {response.text[:200]}...")
                # Return a default response if JSON parsing fails
                return {
                    "error": "Failed to parse API response",
                    "raw_response": response.text[:1000]  # Include truncated raw response
                }
            
        except requests.exceptions.RequestException as e:
            self.logger.error(f"API call failed: {str(e)}")
            raise
        except Exception as e:
            self.logger.error(f"Error in analyze_email_thread: {str(e)}")
            raise

    def format_email_for_api(self, email_content: Dict[str, Any], previous_summary: str = "") -> Dict[str, Any]:
        """Format email content into the structure expected by the API"""
        return {
            "mail_id": email_content.get("id", ""),
            "file_name": [],
            "email": email_content.get("sender", {}).get("emailAddress", {}).get("address", ""),
            "mail_time": email_content.get("receivedDateTime", ""),
            "body_type": "html",
            "mail_body": email_content.get("body", {}).get("content", "").replace("\\n", "\n"),
            "thread_id": email_content.get("conversationId", ""),
            "mail_summary": previous_summary
        }

    def identify_features(self, email_data: Dict[str, Any], features: List[Dict[str, Any]]) -> List[str]:
        """
        Identify which features apply to an email based on conditions
        Returns a list of applicable feature names
        """
        # Format the features for the prompt
        feature_descriptions = "\n".join([
            f"{i+1}. {f['name']}: Description: {f['description']} Condition: {f['condition']} Attributes: {f['attributes']}"
            for i, f in enumerate(features)
        ])
        
        prompt = f"""
        Given the following email:
        
        From: {email_data['from_email']}
        To: {', '.join(email_data['to_email'])}
        Subject: {email_data['subject']}
        Content: {email_data['content']}
        
        And the following features:
        {feature_descriptions}
        
        For each feature, return YES if the email satisfies the feature (based on the description, condition, and attributes), otherwise NO. 
        Return a JSON object with feature names as keys and YES/NO as values.
        """
        
        try:
            response = self.model.generate_content(prompt)
            result = json.loads(response.text)
            return [k for k, v in result.items() if v.strip().upper() == "YES"]
        except Exception as e:
            print(f"Error identifying features: {e}")
            return []
    
    def generate_groundtruth(self, email_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate groundtruth data for an email
        Returns detailed analysis in a structured format
        """
        prompt = f"""
        Analyze the following email and provide a detailed assessment in JSON format:
        
        From: {email_data['from_email']}
        To: {', '.join(email_data['to_email'])}
        Subject: {email_data['subject']}
        Content: {email_data['content']}
        
        Provide a JSON response with these fields:
        1. "email_summarization" - A 1-2 sentence summary of the email content in indirect speech style (third person)
            - Example of correct style: "The sender apologized for the delay and mentioned they will review the materials next week"
        2. "events_summarization" - 

        - Extract all event details mentioned in emails
            - Required event fields: Event name, Date, Time, Property Type, Agent Name, Location
            - Rules:
              * Return empty array ([]) if no events exist
              * Use null for specific fields not mentioned in the email
              * Populate all fields that have explicit values in the email content

            - Example 
            "Events": 

            "Event name": "Visit to Social Security Office",
            "Date": "2025-06-02",
            "Time": "5pm",
            "Property Type": "furnished",
            "Agent Name": "Naveen",
            "Location": "Mountain View SSO"
    

        3. "sentiment_analysis" - The overall tone (red or green based on the email content )

        4. "overall_sentiment_analysis" - The overall tone of the thread as the sum of all the emails the majority wins (positive, negative, neutral)

        5. "category" 
        6. "feature"

            
           FEATURE & CATEGORY IDENTIFICATION:
            - Apply the following classification matrix precisely:
            
            | FEATURE | DESCRIPTION | CATEGORY | CONDITION |
            |---------|-------------|----------|-----------|
            | EMAIL -- DSC First Contact with EE Completed | First email that DSC sends to EE, typically introductory | Initial Service Milestones | IF an email is SENT by DSC for the first time |
            | EMAIL -- EE First Contact with DSC | First email from EE to DSC with availability for initial consultation or relocation status | Initial Service Milestones | IF an email is RECEIVED by DSC for the first time |
            | EMAIL -- Phone Consultation Scheduled | Email confirming specific date/time for first phone consultation | Initial Service Milestones | IF an email is SENT or RECEIVED by DSC and there is NOT already a phone consultation scheduled date in myDW |
            | EMAIL -- Phone Consultation Completed | Email confirming completion of first phone consultation | Initial Service Milestones | IF an email is SENT by DSC and there is NOT already a phone consultation completed date in myDW |
            | No feature | Email does not match any defined features | No category | IF the email does not match any of the above conditions |
            
        
        Ensure the output is valid JSON with these exact field names.
        """
        
        try:
            response = self.model.generate_content(prompt)
            result = json.loads(response.text)
            return result
        except Exception as e:
            print(f"Error generating groundtruth: {e}")
            return {
                "email_summarization": "Error generating summary",
                "events_summarization": None,
                "sentiment_analysis": "neutral",
                "category_analysis": "unknown"
            }
    
    def evaluate_against_groundtruth(self, groundtruth: Dict[str, Any], ai_output: Dict[str, Any]) -> Dict[str, Any]:
        """
        Compare AI output against groundtruth data
        Returns evaluation metrics and comparison results
        """
        # Generate comparison prompt for the AI to evaluate the match
        prompt = f"""
        Compare the groundtruth analysis with the AI output and evaluate the match for each field:
        
        Groundtruth:
        {json.dumps(groundtruth, indent=2)}
        
        AI Output:
        {json.dumps(ai_output, indent=2)}
        
        For each field (email_summarization, events_summarization, sentiment_analysis, category_analysis):
        1. Is the AI output semantically matching the groundtruth? (YES/NO)
        2. Provide a score from 0-100 for the accuracy
        3. Brief explanation of the differences if any and why the AI output is correct or incorrect , mention the patterns and keywords which made you come to this conclusion
        
        Return the evaluation as a JSON object with fields for each category and an overall accuracy score.
        """
        
        try:
            response = self.model.generate_content(prompt)
            evaluation = json.loads(response.text)
            
            # Ensure we have a consistent structure even if the AI response varies
            result = {
                "comparison": evaluation,
                "fields_evaluated": list(groundtruth.keys()),
                "overall_accuracy": evaluation.get("overall_accuracy", 0)
            }
            
            return result
        except Exception as e:
            print(f"Error evaluating against groundtruth: {e}")
            
            # Create a fallback evaluation structure
            fields = ["email_summarization", "events_summarization", "sentiment_analysis", "category_analysis"]
            comparison = {}
            
            for field in fields:
                comparison[field] = {
                    "match": "NO",
                    "score": 0,
                    "explanation": "Error during evaluation"
                }
            
            return {
                "comparison": comparison,
                "fields_evaluated": fields,
                "overall_accuracy": 0
            }
    
    def mock_external_ai_service(self, email_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Mock for an external AI service that would analyze emails
        In a real implementation, this would call your external AI API
        """
        # In reality, this would call your company's AI API
        # For now, we'll generate a simplified response using Gemini
        prompt = f"""
        Analyze the following email and provide a structured assessment:
        
        From: {email_data['from_email']}
        To: {', '.join(email_data['to_email'])}
        Subject: {email_data['subject']}
        Content: {email_data['content']}
        
        Return a JSON object with these fields:
        1. "email_summarization" - A concise summary
        2. "events_summarization" - Events mentioned
        3. "sentiment_analysis" - Overall tone
        4. "category_analysis" - Email type
        """
        
        try:
            response = self.model.generate_content(prompt)
            result = json.loads(response.text)
            return result
        except Exception as e:
            print(f"Error in external AI service: {e}")
            return {
                "email_summarization": "Unable to analyze email",
                "events_summarization": None,
                "sentiment_analysis": "unknown",
                "category_analysis": "unknown"
            }
            
    def use_gemini_for_analysis(self, email_contents: List[Dict[str, Any]], feature_set: str = "real_estate", model: str = "gemini_pro") -> Dict[str, Any]:
        """
        Analyze an email thread and extract relevant features using Gemini
        specifically adapted for real estate and property management communications
        
        Args:
            email_contents: List of email content dictionaries with sender_name, sender_email, content, and sent_time
            feature_set: One of 'real_estate', 'basic', 'advanced'
            model: One of 'gemini_pro', 'gemini_flash'
            
        Returns:
            Dictionary with analysis results
        """
        # Use the specified model
        if model == "gemini_flash":
            self.model = genai.GenerativeModel(model_name="gemini-1.5-flash")
        else:  # default to pro
            self.model = genai.GenerativeModel(model_name="gemini-1.5-pro")
            
        # Look for previous context in the email content
        previous_context = ""
        for email in email_contents:
            if 'previous_context' in email and email['previous_context']:
                previous_context = email['previous_context']
                break
        
        # Format the email thread for the prompt
        email_thread_text = ""
        for i, email in enumerate(email_contents):
            email_thread_text += f"\nEmail {i+1}:\n"
            email_thread_text += f"From: {email.get('sender_name', 'Unknown')} ({email.get('sender_email', 'unknown')})\n"
            if 'recipients' in email:
                email_thread_text += f"To: {', '.join(email.get('recipients', []))}\n"
            email_thread_text += f"Date: {email.get('sent_time', '')}\n"
            email_thread_text += f"Content: {email.get('content', '')}\n"
            email_thread_text += "-" * 40 + "\n"
        
        # Real Estate specific feature extraction prompt
        if feature_set == "real_estate":
            # Include previous summary in the prompt if available
            previous_context_section = ""
            if previous_context:
                previous_context_section = f"""
                ## PREVIOUS EMAIL SUMMARY:
                {previous_context}
                
                Use this previous email summary as context for your analysis. The summary above is from previous emails in this thread.
                """
            
            prompt = f"""
            You are an expert email analyzer specializing in real estate and property management communications.
            Your task is to analyze the following email thread and extract key information according to specific guidelines.
            
            {previous_context_section}
            
            Email Thread:
            {email_thread_text}
            
            ## ANALYSIS REQUIREMENTS:
            
            ### 1. SENTIMENT ANALYSIS:
            - Analyze sentiment based on language tone, keywords, and context
            - Provide TWO sentiment fields:
              * sentiment_analysis: sentiment of the individual email (values: "positive", "negative", or "neutral")
              * overall_sentiment_analysis: sentiment of the entire thread progression (values: "positive", "negative", or "neutral")
            
            ### 2. FEATURE & CATEGORY IDENTIFICATION:
            - Apply the following classification matrix precisely:
            
            | FEATURE | DESCRIPTION | CATEGORY | CONDITION |
            |---------|-------------|----------|-----------|
            | EMAIL -- DSC First Contact with EE Completed | First email that DSC sends to EE, typically introductory | Initial Service Milestones | IF an email is SENT by DSC for the first time |
            | EMAIL -- EE First Contact with DSC | First email from EE to DSC with availability for initial consultation or relocation status | Initial Service Milestones | IF an email is RECEIVED by DSC for the first time |
            | EMAIL -- Phone Consultation Scheduled | Email confirming specific date/time for first phone consultation | Initial Service Milestones | IF an email is SENT or RECEIVED by DSC and there is NOT already a phone consultation scheduled date in myDW |
            | EMAIL -- Phone Consultation Completed | Email confirming completion of first phone consultation | Initial Service Milestones | IF an email is SENT by DSC and there is NOT already a phone consultation completed date in myDW |
            | No feature | Email does not match any defined features | No category | IF the email does not match any of the above conditions |
            
            ### 3. EVENT DETECTION:
            - Extract all event details mentioned in emails
            - Required event fields: Event name, Date, Time, Property Type, Agent Name, Location
            - Rules:
              * Return empty array ([]) if no events exist
              * Use null for specific fields not mentioned in the email
              * Populate all fields that have explicit values in the email content
            
            ### 4. EMAIL SUMMARIZATION:
            - Provide a concise 1-2 sentence summary of the email content
            - Focus on key details like property information, scheduling, or action items
            
            ## RESPONSE FORMAT:
            Return your analysis as a valid JSON object with these fields:
            
            ```json
            {{
                "sentiment_analysis": "positive|negative|neutral",
                "overall_sentiment_analysis": "positive|negative|neutral", 
                "feature": "EMAIL -- DSC First Contact with EE Completed|EMAIL -- EE First Contact with DSC|EMAIL -- Phone Consultation Scheduled|EMAIL -- Phone Consultation Completed|No feature",
                "category": "Initial Service Milestones|No category",
                "Events": [
                    {{
                        "Event name": "string or null",
                        "Date": "string or null",
                        "Time": "string or null",
                        "Property Type": "string or null",
                        "Agent Name": "string or null",
                        "Location": "string or null"
                    }}
                ],
                "Summary": "1-2 sentence summary of the email"
            }}
            ```
            
            Ensure your analysis precisely follows the required format and classification matrix.
            """
        else:
            # For other feature sets, use the original prompt but include previous context
            features_to_extract = ["email_summarization", "sentiment_analysis", "overall_sentiment_analysis", 
                                  "feature", "category", "Events"]
            
            if feature_set == "advanced":
                features_to_extract.extend(["key_entities", "response_suggestion", "priority_level"])
        
            # Generate the feature extraction prompt
            features_list = ", ".join(features_to_extract)
            
            # Include previous summary in the prompt if available
            previous_context_section = ""
            if previous_context:
                previous_context_section = f"""
                Previous Email Summary Context:
                {previous_context}
                
                Use this previous email summary as context for your analysis.
                """
            
            prompt = f"""
            Analyze the following email thread and extract these features: {features_list}
            
            {previous_context_section}
            
            Email Thread:
            {email_thread_text}
            
            For each feature, provide a detailed analysis. For boolean features (is_urgent, is_action_required), 
            return true/false and a confidence score between 0.0 and 1.0.
            
            Return your analysis as a valid JSON object with the feature names as keys.
            Include confidence scores (_confidence suffix) for each feature when applicable.
            """
        
        try:
            response = self.model.generate_content(prompt)
            content = response.text
            
            # Try to extract JSON from the response
            start_idx = content.find('{')
            end_idx = content.rfind('}')
            if start_idx != -1 and end_idx != -1:
                json_content = content[start_idx:end_idx+1]
                result = json.loads(json_content)
            else:
                # Fall back to a simpler prompt if JSON extraction fails
                # Include previous context in this prompt too
                previous_context_simple = f"Previous email summary: {previous_context}\n\n" if previous_context else ""
                
                simple_prompt = f"""
                {previous_context_simple}
                Analyze the following email thread:
                
                {email_thread_text}
                
                Provide ONLY a valid JSON object with these fields:
                * sentiment_analysis: sentiment of the email (positive, negative, neutral)
                * overall_sentiment_analysis: sentiment of the thread (positive, negative, neutral)
                * feature: feature classification based on the provided matrix
                * category: appropriate category from the matrix
                * Events: array of event details (can be empty)
                * Summary: 1-2 sentence summary
                
                Ensure the response is a properly formatted JSON object with no additional text.
                """
                response = self.model.generate_content(simple_prompt)
                content = response.text
                
                start_idx = content.find('{')
                end_idx = content.rfind('}')
                if start_idx != -1 and end_idx != -1:
                    json_content = content[start_idx:end_idx+1]
                    result = json.loads(json_content)
                else:
                    # If we still can't get proper JSON, return a basic structure
                    result = {
                        "sentiment_analysis": "neutral",
                        "overall_sentiment_analysis": "neutral",
                        "feature": "No feature",
                        "category": "No category",
                        "Events": [],
                        "Summary": "Unable to analyze email thread"
                    }
            
            # Add metadata about the analysis
            result["_feature_set"] = feature_set
            result["_model_used"] = model
            result["_thread_length"] = len(email_contents)
            result["_analysis_timestamp"] = str(datetime.now())
            result["_previous_context_used"] = bool(previous_context)
            
            return result
            
        except Exception as e:
            print(f"Error analyzing email thread: {str(e)}")
            
            # Return a basic analysis structure as fallback
            return {
                "sentiment_analysis": "neutral",
                "overall_sentiment_analysis": "neutral",
                "feature": "No feature",
                "category": "No category",
                "Events": [],
                "Summary": "Error occurred during analysis",
                "_error": str(e),
                "_feature_set": feature_set,
                "_model_used": model,
                "_thread_length": len(email_contents),
                "_previous_context_used": bool(previous_context)
            }