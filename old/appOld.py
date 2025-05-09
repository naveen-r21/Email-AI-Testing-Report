from dotenv import load_dotenv # Add this import
load_dotenv() # Add this call to load .env file

import streamlit as st
import os
import json
import pandas as pd
from datetime import datetime, timedelta
import time
import re
from io import BytesIO
import requests
import urllib.parse
import traceback
from typing import Dict, List, Any, Optional, Tuple, Union
from bs4 import BeautifulSoup
from thefuzz import fuzz
import altair as alt
import secrets
import google.generativeai as genai
from evaluation_engine import evaluate_with_llm, calculate_similarity, generate_fallback_metrics
import logging
import sys # Add this at the top of app.py for df.info() printing
from graph_api_client import GraphAPIClient # Added this import
from ai_service import AIService # Added this import

# Helper function to get clean subject by removing prefixes like Re:, Fw:, etc. from email subjects
def get_clean_subject(subject):
    """Remove prefixes like Re:, Fw:, etc. from email subjects"""
    if not subject:
        return "No Subject"
    return re.sub(r'^(?:(?:Re|RE|Fw|FW|Fwd|FWD):\s*)+', '', subject).strip()

# Helper function to render metrics in the desired table format
def render_metric_display_table(metric_dict, metric_name_for_title):
    import pandas as pd 
    import streamlit as st 
    import json # Add json import for handling event data
    
    # >>> Debug print removed here <<<
    
    st.subheader(metric_name_for_title.replace("_", " ").title())

    table_data = []
    ai_value = metric_dict.get("AI Value", "N/A")
    gt_value = metric_dict.get("Ground Truth", "N/A")

    # --- START EVENT HANDLING ---
    # If it's event data and the value is a list or dict, convert to JSON string for display
    if "event" in metric_name_for_title.lower():
        if isinstance(ai_value, (list, dict)):
            try:
                ai_value = json.dumps(ai_value, indent=2)
            except Exception:
                ai_value = str(ai_value) # Fallback to string conversion
        if isinstance(gt_value, (list, dict)):
            try:
                gt_value = json.dumps(gt_value, indent=2)
            except Exception:
                gt_value = str(gt_value) # Fallback to string conversion
    # --- END EVENT HANDLING ---
            
    # Always add AI Value and Ground Truth (potentially formatted)
    table_data.append({"Metric": "AI Value", "Value": ai_value})
    table_data.append({"Metric": "Ground Truth", "Value": gt_value})

    # Add Similarity Percentage ONLY for Summary metric
    if metric_name_for_title.lower() == "summary":
        similarity_val = metric_dict.get("Similarity Percentage", metric_dict.get("Similarity")) # Check both keys
        # Basic formatting to ensure it looks like a percentage string
        if isinstance(similarity_val, (int, float)):
             similarity_val = f"{similarity_val:.0%}"
        elif isinstance(similarity_val, str) and "%" not in similarity_val:
             try:
                 similarity_float = float(similarity_val)
                 similarity_val = f"{similarity_float:.0%}"
             except ValueError:
                 pass # Keep string as is if conversion fails
        table_data.append({"Metric": "Similarity Percentage", "Value": similarity_val if similarity_val is not None else "N/A"})

    # Add Pass/Fail Status
    table_data.append({"Metric": "Pass/Fail", "Value": metric_dict.get("Status", "N/A")})
    
    # Add Ground Truth Explanation (Mandatory - should not be N/A from LLM)
    table_data.append({"Metric": "Ground Truth Explanation", "Value": metric_dict.get("Ground Truth Explanation", "Explanation missing from LLM")})
    
    # Add Pass/Fail or % Explanation
    explanation_key = "% Explanation" if metric_name_for_title.lower() == "summary" else "Pass/Fail Explanation"
    explanation_value = metric_dict.get(explanation_key, metric_dict.get("Pass/Fail Explanation", metric_dict.get("Comparison Explanation", "Explanation missing from LLM")))
    table_data.append({"Metric": explanation_key, "Value": explanation_value})
    
    df = pd.DataFrame(table_data)
    print(f"DEBUG render_metric_display_table: Constructed DataFrame for '{metric_name_for_title}':\n{df.to_string()}")
    
    st.table(df.set_index("Metric"))

# Function to convert metrics data to Excel format
def convert_metrics_to_excel(individual_results):
    """Converts the list of individual results into an Excel file buffer with two sheets."""
    import pandas as pd 
    from io import BytesIO
    import re

    # --- Sheet 1: Detailed Evaluation Metrics --- 
    all_metrics_data = []
    for idx, result in enumerate(individual_results):
        email_subject = result.get('email', {}).get('subject', f'Email {idx+1} - No Subject')
        metrics = result.get('metrics', [])
        for metric in metrics:
            row_data = {
                "Email Index": idx + 1,
                "Email Subject": get_clean_subject(email_subject),
                "Metric": metric.get("Metric", metric.get("Field", "Unknown")),
                "Status": metric.get("Status", "N/A"),
                "Pass/Fail Explanation": metric.get("Pass/Fail Explanation", metric.get("Comparison Explanation", "N/A")),
                "Similarity Percentage": metric.get("Similarity Percentage", metric.get("Similarity", "N/A")),
                "% Explanation": metric.get("% Explanation", "N/A"),
                "Ground Truth": metric.get("Ground Truth", "N/A"),
                "Ground Truth Explanation": metric.get("Ground Truth Explanation", metric.get("GT Explanation", "N/A")),
                "AI Value": metric.get("AI Value", "N/A"),
                "Individual Email Review Points": metric.get("individual_email_review_points", "N/A") # New field
            }
            if row_data["Metric"].lower() != 'summary':
                row_data["Similarity Percentage"] = ""
                row_data["% Explanation"] = ""
            if row_data["Metric"].lower() == 'summary':
                 row_data["Pass/Fail Explanation"] = "" # Use % Explanation instead
                 
            all_metrics_data.append(row_data)

    df_metrics = pd.DataFrame(all_metrics_data if all_metrics_data else [{"Info": "No detailed metrics."}])
    if "Email Index" in df_metrics.columns: # Check if we have actual metrics
        cols_order = [
            "Email Index", "Email Subject", "Metric", "Status",
            "Pass/Fail Explanation", "Similarity Percentage", "% Explanation",
            "Ground Truth", "Ground Truth Explanation", "AI Value",
            "Individual Email Review Points" # New field
        ]
        cols_to_use = [col for col in cols_order if col in df_metrics.columns]
        df_metrics = df_metrics[cols_to_use]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_metrics.to_excel(writer, index=False, sheet_name='Evaluation Metrics')
        worksheet_metrics = writer.sheets['Evaluation Metrics']
        text_wrap_format = writer.book.add_format({'text_wrap': True, 'valign': 'top'})
        explanation_cols = ['Pass/Fail Explanation', '% Explanation', 'Ground Truth Explanation', 'Individual Email Review Points'] # Added new col
        for col_num, column_title in enumerate(df_metrics.columns):
            max_len = max(df_metrics[column_title].astype(str).map(len).max(), len(column_title)) + 2
            if column_title in explanation_cols:
                worksheet_metrics.set_column(col_num, col_num, min(max_len, 60), text_wrap_format) # Increased cap for review points
            else:
                worksheet_metrics.set_column(col_num, col_num, min(max_len, 30))

        # --- Sheet 2: Overall Thread Review --- 
        overall_review_text = generate_overall_thread_review(individual_results) # This is the 5-point text

        # Calculate Dashboard Metrics (these should ideally be passed in or recalculated cleanly)
        # For now, recalculating here for simplicity, mirroring the UI dashboard logic
        total_mails_excel = len(individual_results)
        total_fields_validated_excel = 0
        total_pass_excel = 0
        for res_excel in individual_results:
            metrics_for_email_excel = res_excel.get('metrics', [])
            for metric_item_excel in metrics_for_email_excel:
                total_fields_validated_excel += 1
                if metric_item_excel.get("Status") == "Pass":
                    total_pass_excel += 1
                if "event" in metric_item_excel.get("Metric", metric_item_excel.get("Field", "")).lower():
                    ai_event_val_excel = metric_item_excel.get("AI Value")
                    try:
                        if isinstance(ai_event_val_excel, str):
                            parsed_ai_events_excel = json.loads(ai_event_val_excel)
                            if isinstance(parsed_ai_events_excel, list) and len(parsed_ai_events_excel) > 0:
                                total_fields_validated_excel += (len(parsed_ai_events_excel[0]) -1) * len(parsed_ai_events_excel)
                    except: pass
        accuracy_percentage_excel = (total_pass_excel / total_fields_validated_excel * 100) if total_fields_validated_excel > 0 else 0
        
        # Prepare data for Overall Review sheet
        review_sheet_data = []
        review_sheet_data.append({"Category": "Dashboard Metric", "Detail": "Total Mails Processed", "Value": total_mails_excel})
        review_sheet_data.append({"Category": "Dashboard Metric", "Detail": "Total Fields Validated", "Value": total_fields_validated_excel})
        review_sheet_data.append({"Category": "Dashboard Metric", "Detail": "Total Fields Passed", "Value": total_pass_excel})
        review_sheet_data.append({"Category": "Dashboard Metric", "Detail": "Overall Accuracy", "Value": f"{accuracy_percentage_excel:.2f}%"})
        review_sheet_data.append({"Category": "-", "Detail": "-", "Value": "-"}) # Separator row

        # Add the 5-point review from Gemini
        review_sheet_data.append({"Category": "Gemini Overall Review", "Detail": "(Generated by AI)", "Value": ""})
        for i, line in enumerate(overall_review_text.split('\n')):
            if line.strip(): # Add non-empty lines
                 review_sheet_data.append({"Category": f"Point {i+1}" if line.strip()[0].isdigit() else "Review Text", "Detail": line.strip(), "Value": ""})

        df_review = pd.DataFrame(review_sheet_data)
        
        df_review.to_excel(writer, index=False, sheet_name='Overall Review')
        worksheet_review = writer.sheets['Overall Review']
        text_wrap_format = writer.book.add_format({'text_wrap': True, 'valign': 'top'})
        worksheet_review.set_column(0, 0, 25, text_wrap_format) # Category
        worksheet_review.set_column(1, 1, 70, text_wrap_format) # Detail
        worksheet_review.set_column(2, 2, 20, text_wrap_format) # Value (for dashboard metrics)
            
    processed_data = output.getvalue()
    return processed_data

# Function to generate overall review
def generate_overall_thread_review(individual_results: List[Dict[str, Any]]) -> str:
    """Generates an overall review of the thread using Gemini based on individual email metrics."""
    if not individual_results:
        return "No results available to generate an overall review."

    # Compile findings from all emails
    compiled_findings = []
    for idx, result in enumerate(individual_results):
        email_subject = result.get('email', {}).get('subject', f'Email {idx+1}')
        metrics = result.get('metrics', [])
        findings_for_email = [f"Email {idx+1} ({email_subject}):"] 
        for m in metrics:
            metric_name = m.get("Metric", m.get("Field", "Unknown"))
            status = m.get("Status", "N/A")
            pf_explanation = m.get("Pass/Fail Explanation", m.get("Comparison Explanation", ""))
            findings_for_email.append(f"  - {metric_name}: {status}. Explanation: {pf_explanation[:150]}...") # Truncate explanation
        compiled_findings.append("\n".join(findings_for_email))
        
    full_findings_summary = "\n\n".join(compiled_findings)

    # Call Gemini for the overall review
    try:
        # Use API key from session state
        api_key = st.session_state.get("gemini_api_key")
        if not api_key:
            return "Error: Gemini API Key not configured in settings."
            
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')

        prompt = f"""
        Analyze the following summary of AI evaluation results for an entire email thread. 
        Provide an overall review, highlighting what the AI did well and where it needs improvement across the whole thread.

        Evaluation Summary:
        {full_findings_summary}

        TASK:
        Generate a concise, 5-point overall review of the AI's performance on this thread.
        Focus on consistent patterns (good or bad) across the emails.
        
        Output Format (Exactly 5 points):
        1. [Point 1: Observation about overall performance - e.g., Sentiment accuracy]
        2. [Point 2: Observation about overall performance - e.g., Feature/Category identification]
        3. [Point 3: Observation about overall performance - e.g., Summary quality/consistency]
        4. [Point 4: Strength observed across the thread]
        5. [Point 5: Area needing improvement across the thread]
        """

        response = model.generate_content(prompt)
        return response.text.strip()

    except Exception as e:
        print(f"Error generating overall review with Gemini: {str(e)}")
        return f"Error generating overall review: {str(e)}"

# Define the fetch_threads function to properly fetch and filter email threads
def fetch_threads(graph_client, user_email, from_email=None, to_email=None, subject_contains=None, date_from: Optional[datetime.date] = None, date_to: Optional[datetime.date] = None):
    """
    Fetch email threads and apply filters
    
    Args:
        graph_client: The GraphAPIClient instance
        user_email: User's email address
        from_email: Filter by sender email
        to_email: Filter by recipient email
        subject_contains: Filter by subject text
        date_from: Optional start date for filtering
        date_to: Optional end date for filtering
        
    Returns:
        tuple: (list of threads, error message if any)
    """
    try:
        print(f"Fetching threads for user: {user_email}")
        print(f"Filters - From: {from_email}, To: {to_email}, Subject: {subject_contains}, DateFrom: {date_from}, DateTo: {date_to}")
        
        # Use the built-in API client methods to get threads
        thread_list = graph_client.group_emails_by_subject(user_email, count=100)
        
        if not thread_list:
            return [], "No email threads found"
        
        print(f"Found {len(thread_list)} threads before filtering")
        
        # Apply filters if provided
        filtered_thread_list = []
        for thread in thread_list:
            # Default to including the thread
            include_thread = True
            
            # Get the first message if available
            first_message = thread.get('first_message', {})
            latest_message = thread.get('latest_message', {})
            
            # Apply from_email filter if provided
            if from_email and include_thread:
                # Check if we have the first_message object
                if isinstance(first_message, dict):
                    sender = first_message.get('from', {}).get('emailAddress', {}).get('address', '')
                    if from_email.lower() not in sender.lower():
                        # Also check the latest message
                        if isinstance(latest_message, dict):
                            sender = latest_message.get('from', {}).get('emailAddress', {}).get('address', '')
                            if from_email.lower() not in sender.lower():
                                include_thread = False
                                print(f"Thread excluded by from filter: {thread.get('subject', 'No Subject')}")
                        else:
                            include_thread = False
                else:
                    # If first_message is not available, check participants
                    participants = thread.get('participants', [])
                    if not any(from_email.lower() in p.lower() for p in participants):
                        include_thread = False
                        print(f"Thread excluded by from filter (participants check): {thread.get('subject', 'No Subject')}")
            
            # Apply to_email filter if provided
            if to_email and include_thread:
                if isinstance(first_message, dict):
                    recipients = first_message.get('toRecipients', [])
                    recipient_emails = []
                    for r in recipients:
                        if isinstance(r, dict) and 'emailAddress' in r:
                            recipient_emails.append(r.get('emailAddress', {}).get('address', ''))
                    
                    if not any(to_email.lower() in email.lower() for email in recipient_emails):
                        include_thread = False
                        print(f"Thread excluded by to filter: {thread.get('subject', 'No Subject')}")
            
            # Apply subject filter if provided
            if subject_contains and include_thread:
                subject = thread.get('subject', '')
                if subject_contains.lower() not in subject.lower():
                    include_thread = False
                    print(f"Thread excluded by subject filter: {thread.get('subject', 'No Subject')}")
            
            # Apply date filters if provided
            if include_thread and (date_from or date_to):
                thread_date_str = thread.get('latest_message_date')
                if thread_date_str:
                    try:
                        # Ensure datetime is imported if not already at the top of the file
                        # from datetime import datetime # This should be at the top of app.py
                        thread_date = datetime.fromisoformat(thread_date_str.replace('Z', '')).date()
                        
                        if date_from and thread_date < date_from:
                            include_thread = False
                            print(f"Thread excluded by date_from filter: {thread.get('subject', 'No Subject')} (Date: {thread_date})")
                        
                        if include_thread and date_to and thread_date > date_to:
                            include_thread = False
                            print(f"Thread excluded by date_to filter: {thread.get('subject', 'No Subject')} (Date: {thread_date})")
                            
                    except ValueError:
                        print(f"Warning: Could not parse date '{thread_date_str}' for thread {thread.get('subject', 'No Subject')}. Skipping date filter for this thread.")
                else:
                    print(f"Warning: No 'latest_message_date' found for thread {thread.get('subject', 'No Subject')}. Skipping date filter for this thread.")

            # Add thread to filtered list if it passed all filters
            if include_thread:
                filtered_thread_list.append(thread)
        
        print(f"Found {len(filtered_thread_list)} threads after filtering")
        return filtered_thread_list, None
        
    except Exception as e:
        print(f"Error in fetch_threads: {str(e)}")
        print(traceback.format_exc())
        return None, f"Error fetching threads: {str(e)}"

# Initialize session state variables if they don't exist
def initialize_session_state():
    """Initialize or update session state variables."""
    
    # --- API Credentials & Settings ---
    # For each setting, try to load from environment, then use session state if already set by user
    # This ensures that if a user sets a value in the UI, it persists for their session, 
    # overriding any environment variable for that session only.
    # If nothing is in session_state (e.g., first run or after clearing cache), 
    # it attempts to load from os.getenv().
    
    defaults = {
        "ms_graph_client_id": os.getenv("MS_GRAPH_CLIENT_ID", ""),
        "ms_graph_client_secret": os.getenv("MS_GRAPH_CLIENT_SECRET", ""),
        "ms_graph_tenant_id": os.getenv("MS_GRAPH_TENANT_ID", ""),
        "user_email": os.getenv("USER_EMAIL", ""), # User whose emails to fetch
        "gemini_api_key": os.getenv("GEMINI_API_KEY", ""),
        "dwellworks_api_endpoint": os.getenv("DWELLWORKS_API_ENDPOINT", ""), # User updated this default
        # "dwellworks_api_key": os.getenv("DWELLWORKS_API_KEY", ""), # REMOVED
        
        # --- UI State & Data ---
        'current_page': 'main_app',
        'selected_email_id': None,
        'selected_thread_id': None, # Added for thread selection
        'email_thread_data': None, # Store fetched thread data
        'individual_results': [],
        'overall_summary': None,
        'user_query': "",
        'last_search_params': None,
        'show_retrieval_details': False, 
        'show_processing_details': False,
        'show_evaluation_details': False,
        'show_excel_summary': False,
        'last_processed_ids': [],
        'metrics_df': None,
        'evaluation_results': None,
        'search_results': None,
        'processing_log': [],
        'error_messages': [],
        'info_messages': [],
        'active_thread_emails': None, # Store emails of the active thread being viewed
        'active_thread_subject': None, # Store subject of the active thread
        'current_search_query_display': "No active search.",
        'upload_key_counter': 0,  # For resetting file uploader
        'email_source_option': 'Upload Email Excel', # Default to Excel upload
        'excel_file_processed': False,
        'graph_api_results_processed': False,
        'ai_service_instance': None,
        'graph_client_instance': None,
        'raw_email_content_for_display': None,
        'parsed_email_content_for_display': None,
        'overall_thread_review': None, # Added for storing the 5-point review
        'system_prompt_template': "", # Initialize system prompt template
        'user_prompt_template': "",   # Initialize user prompt template
        'ground_truth_template': ""  # Initialize ground truth template
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

# Ensure this is called ONCE at the beginning of the script
if 'session_initialized' not in st.session_state:
    initialize_session_state()
    st.session_state.session_initialized = True

# --- Initialize API Clients (GraphAPICLient and AIService) ---
# These clients will now use credentials from st.session_state, 
# which are populated by initialize_session_state from env vars or kept blank if not set.
# They are re-initialized if relevant settings change (e.g., after user clicks "Save Settings").

def get_graph_client():
    if st.session_state.ms_graph_client_id and \
       st.session_state.ms_graph_client_secret and \
       st.session_state.ms_graph_tenant_id:
        try:
            return GraphAPIClient(
                client_id=st.session_state.ms_graph_client_id,
                client_secret=st.session_state.ms_graph_client_secret,
                tenant_id=st.session_state.ms_graph_tenant_id
            )
        except Exception as e:
            st.error(f"Error initializing Graph API Client: {e}")
            return None
    return None

def get_ai_service():
    # AIService now only needs dwellworks_api_endpoint and gemini_api_key
    if st.session_state.dwellworks_api_endpoint and st.session_state.gemini_api_key:
        try:
            return AIService(
                dwellworks_api_endpoint=st.session_state.dwellworks_api_endpoint,
                # dwellworks_api_key=st.session_state.dwellworks_api_key, # REMOVED
                gemini_api_key=st.session_state.gemini_api_key
            )
        except Exception as e:
            st.error(f"Error initializing AI Service: {e}")
            return None
    return None

# Store client instances in session state to avoid re-creation on every rerun unless settings change
if 'graph_client_instance' not in st.session_state or \
   st.session_state.get('graph_client_instance_config') != (
       st.session_state.ms_graph_client_id, 
       st.session_state.ms_graph_client_secret, 
       st.session_state.ms_graph_tenant_id):
    st.session_state.graph_client_instance = get_graph_client()
    st.session_state.graph_client_instance_config = (
        st.session_state.ms_graph_client_id, 
        st.session_state.ms_graph_client_secret, 
        st.session_state.ms_graph_tenant_id
    )

if 'ai_service_instance' not in st.session_state or \
    st.session_state.get('ai_service_instance_config') != (
        st.session_state.dwellworks_api_endpoint, 
        # st.session_state.dwellworks_api_key, # REMOVED
        st.session_state.gemini_api_key):
    st.session_state.ai_service_instance = get_ai_service()
    st.session_state.ai_service_instance_config = (
        st.session_state.dwellworks_api_endpoint, 
        # st.session_state.dwellworks_api_key, # REMOVED
        st.session_state.gemini_api_key
    )

# --- Sidebar Setup ---
st.sidebar.title("Email Analysis Pipeline")

# --- API Settings Section ---
with st.sidebar.expander("⚙️ API Settings", expanded=False):
    st.write("Configure API credentials. Values from environment variables are used as defaults if set.")
    
    # MS Graph API Credentials
    # The st.text_input widgets will use their current value (if user typed something)
    # or fall back to st.session_state which was initialized from os.getenv()
    ms_graph_client_id_input = st.text_input(
        "MS Graph Client ID", 
        value=st.session_state.ms_graph_client_id, 
        key="ms_graph_client_id_ui", # Unique key for UI widget
        type="password",
        help="Your Azure App Registration Client ID."
    )
    ms_graph_client_secret_input = st.text_input(
        "MS Graph Client Secret", 
        value=st.session_state.ms_graph_client_secret, 
        key="ms_graph_client_secret_ui",
        type="password",
        help="Your Azure App Registration Client Secret."
    )
    ms_graph_tenant_id_input = st.text_input(
        "MS Graph Tenant ID", 
        value=st.session_state.ms_graph_tenant_id, 
        key="ms_graph_tenant_id_ui",
        type="password",
        help="Your Azure AD Tenant ID."
    )
    user_email_for_graph_input = st.text_input(
        "User Email (to fetch emails for)", 
        value=st.session_state.user_email, 
        key="user_email_ui",
        help="The email address of the user whose mailbox will be accessed (e.g., your.email@example.com)."
    )
    
    # Gemini API Key
    gemini_api_key_input = st.text_input(
        "Gemini API Key", 
        value=st.session_state.gemini_api_key, 
        key="gemini_api_key_ui",
        type="password",
        help="Your Google AI Studio Gemini API Key."
    )

    # Dwellworks API (Optional, can be shown if needed)
    # dwellworks_api_endpoint_input = st.text_input(
    #     "Dwellworks API Endpoint", 
    #     value=st.session_state.dwellworks_api_endpoint, 
    #     key="dwellworks_api_endpoint_ui"
    # )
    # dwellworks_api_key_input = st.text_input(
    #     "Dwellworks API Key", 
    #     value=st.session_state.dwellworks_api_key, 
    #     key="dwellworks_api_key_ui", 
    #     type="password"
    # )

    if st.button("Save API Settings"):
        # When button is clicked, update session_state from the UI input fields
        st.session_state.ms_graph_client_id = ms_graph_client_id_input
        st.session_state.ms_graph_client_secret = ms_graph_client_secret_input
        st.session_state.ms_graph_tenant_id = ms_graph_tenant_id_input
        st.session_state.user_email = user_email_for_graph_input
        st.session_state.gemini_api_key = gemini_api_key_input
        # st.session_state.dwellworks_api_endpoint = dwellworks_api_endpoint_input
        # st.session_state.dwellworks_api_key = dwellworks_api_key_input
        
        # Trigger re-initialization of clients with new settings
        st.session_state.graph_client_instance = get_graph_client()
        st.session_state.graph_client_instance_config = (
            st.session_state.ms_graph_client_id, 
            st.session_state.ms_graph_client_secret, 
            st.session_state.ms_graph_tenant_id
        )
        st.session_state.ai_service_instance = get_ai_service()
        st.session_state.ai_service_instance_config = (
            st.session_state.dwellworks_api_endpoint,
            # st.session_state.dwellworks_api_key, # REMOVED
            st.session_state.gemini_api_key
        )

        st.success("API Settings Saved and services re-initialized!")
        st.rerun() # Rerun to reflect changes immediately

# --- How to Use This Tool Section ---
with st.sidebar.expander("📚 How to Use This Tool", expanded=False):
    st.markdown("""
    ### Step 1: Email Fetching Tab
    - Enter your Microsoft Graph API credentials in the sidebar
    - Use the filter options to narrow down emails if needed
    - Click 'Fetch Email Threads' to retrieve emails
    - Select an email thread from the dropdown to load the conversation
    
    ### Step 2: Results & Reports Tab
    - View the complete analysis results
    - Compare AI-generated features with groundtruth
    - Download reports in Excel or JSON format
    """)

# ------------------------------------------------------------
# Main App Layout
# ------------------------------------------------------------

# Main App Layout - Ensure consistent tab naming
# Create tabs for different sections
tab1, tab2 = st.tabs([
    "Email Fetching",
    "Detailed Evaluation"
])

# ------------------------------------------------------------
# Tab 1: Email Fetching
# ------------------------------------------------------------

with tab1:
    # Remove the Email Fetching header
    # st.header("Email Fetching")
    
    # Initialize session state for this tab if needed
    if 'threads' not in st.session_state:
        st.session_state.threads = []
    
    st.subheader("Select Email Source")
    email_source = st.radio(
        "Choose how to load emails:",
        ("Fetch from Outlook (Graph API)", "Upload Email Excel"),
        index=1,  # Set default to "Upload Email Excel"
        key="email_source_selection"
    )
    st.markdown("---")

    if email_source == "Fetch from Outlook (Graph API)":
        # Filters section
        st.markdown('<div class="filter-container">', unsafe_allow_html=True)
        st.subheader("Email Filters (Graph API)")
    
        col1, col2 = st.columns(2)
    
        with col1:
            from_filter = st.text_input("From (Sender Email)", value="nraman@dwellworks.com")
            to_filter = st.text_input("To (Recipient Email)")
    
        with col2:
            subject_filter = st.text_input("Subject Contains")
            use_date_from = st.checkbox("Enable From Date filter")
            if use_date_from:
                date_from = st.date_input("From Date", value=datetime.now() - timedelta(days=7))
            else:
                date_from = None
            use_date_to = st.checkbox("Enable To Date filter")
            if use_date_to:
                date_to = st.date_input("To Date", value=datetime.now())
            else:
                date_to = None

        # Fetch threads button
        fetch_button = st.button("Fetch Email Threads", key="fetch_threads_btn")
    
        # Close the filter-container div
        st.markdown('</div>', unsafe_allow_html=True)
    
        if fetch_button:
            with st.spinner("Fetching email threads..."):
                try:
                    # Get the GraphAPIClient instance from session state
                    graph_client = st.session_state.graph_client_instance
                    user_email_for_fetching = st.session_state.user_email # Get user_email from session_state

                    if not graph_client:
                        st.error("Graph API client is not initialized. Please check API Settings.")
                    elif not user_email_for_fetching:
                        st.error("User Email for fetching is not configured. Please check API Settings.")
                    else:
                        access_token = graph_client.get_access_token() 
                        if not access_token:
                            st.error('Failed to get access token. Check your credentials and API Settings.')
                        else:
                            # Get and filter threads
                            thread_list, error = fetch_threads(
                                graph_client,
                                user_email_for_fetching, # Use user_email from session_state
                                from_email=from_filter if from_filter else None,
                                to_email=to_filter if to_filter else None,
                                subject_contains=subject_filter if subject_filter else None,
                                date_from=date_from if use_date_from else None,
                                date_to=date_to if use_date_to else None
                            )
                        
                            if error:
                                st.error(error)
                            elif thread_list:
                                # Store in session state
                                st.session_state.threads = thread_list
                                total_messages = sum(t['message_count'] for t in thread_list)
                                st.success(f"Found {len(thread_list)} email threads containing {total_messages} total messages!")
                            else:
                                st.warning("No email threads found matching the specified filters.")
                        
                except Exception as e:
                    st.error(f"Error: {str(e)}")
                    print(f"Error details: {traceback.format_exc()}")

        # Display threads in selectbox for user selection
        thread_options = []
        thread_labels = {}
    
        if 'threads' in st.session_state and st.session_state.threads:
            for thread_idx, thread in enumerate(st.session_state.threads): 
                label = f"{thread['subject']} (Count: {thread['message_count']} emails) - ID: {thread['thread_id'][:8]}"
                thread_labels[label] = thread
                thread_options.append(label)
            
            selected_thread_label = st.selectbox(
                "Select Email Thread to Process:",
                options=thread_options,
                index=0 if thread_options else None,
                key="selected_thread_selectbox" 
            )
            
            if not selected_thread_label:
                st.info("No threads available. Please fetch emails first.")
            else:
                selected_thread = thread_labels[selected_thread_label]
            
                st.markdown("### Thread Information")
                col1_info, col2_info = st.columns(2)
            
                with col1_info:
                    st.markdown(f"**Subject:** {selected_thread['subject']}")
                    st.markdown(f"**Number of Emails:** {selected_thread['message_count']}")
            
                with col2_info:
                    st.markdown(f"**Thread ID:** {selected_thread['thread_id']}")
                    st.markdown(f"**Latest Activity:** {selected_thread['latest_message_date']}")
            
                if selected_thread.get('messages'):
                    with st.expander("🔍 View Email IDs in Thread Group", expanded=False):
                        # Displaying the list of message IDs from the subject-grouped thread
                        if isinstance(selected_thread['messages'], list) and all(isinstance(item, str) for item in selected_thread['messages']):
                            st.json(selected_thread['messages'][:5]) # Show first 5 message IDs
                        else:
                            # Fallback if the structure is not a simple list of strings (as per screenshot anomaly)
                            st.warning("Email message data in this thread group is not in the expected list-of-IDs format.")
                            st.json(selected_thread['messages'][:2]) # Show raw structure for debugging

                if st.button("Process Thread", type="primary", key="process_thread_api_btn"):
                    with st.spinner("Loading and analyzing email thread..."):
                        try:
                            # Get the GraphAPIClient instance and user_email from session state
                            graph_client = st.session_state.graph_client_instance
                            user_email_for_processing = st.session_state.user_email

                            if not graph_client:
                                st.error("Graph API client is not initialized. Please check API Settings.")
                            elif not user_email_for_processing:
                                st.error("User Email for processing is not configured. Please check API Settings.")
                            else:
                                access_token = graph_client.get_access_token()
                        
                                if not access_token:
                                    st.error('Failed to get access token. Check your credentials and API Settings.')
                                else:
                                    st.session_state.individual_results = []
                                    st.session_state.has_results = False
                                    st.session_state.last_processed_source = "graph_api" # Flag source
                                
                                    message_ids_to_fetch = selected_thread.get('messages', [])
                                    fetched_email_objects = [] 

                                    progress_bar_api = st.progress(0) 
                                    error_placeholder_api = st.empty() 
                                    has_rate_limit_error_api = False 
                                    results_api = [] 
                                    previous_summary_api = "" 

                                    if not message_ids_to_fetch:
                                        st.warning("Selected thread contains no message IDs to process.")
                                        progress_bar_api.empty()
                                    else:
                                        st.write(f"Fetching {len(message_ids_to_fetch)} email(s) for the thread...")
                                        for i, msg_id in enumerate(message_ids_to_fetch):
                                            progress_bar_api.progress( (i + 1) / len(message_ids_to_fetch), text=f"Fetching email {i+1}/{len(message_ids_to_fetch)} (ID: {msg_id[:20]}...)")
                                            try:
                                                email_obj = graph_client._get_email_with_body(user_email_for_processing, msg_id)
                                                if email_obj and not email_obj.get("error"): 
                                                    fetched_email_objects.append(email_obj)
                                                else:
                                                    error_text = email_obj.get('message', f'Failed to fetch email ID {msg_id}') if isinstance(email_obj, dict) else f'Failed to fetch email ID {msg_id}'
                                                    error_placeholder_api.warning(error_text)
                                                    print(f"Error fetching email {msg_id}: {email_obj}") 
                                            except Exception as e_fetch_indiv:
                                                error_placeholder_api.warning(f"Exception fetching email ID {msg_id}: {str(e_fetch_indiv)}")
                                                print(f"Exception fetching email {msg_id}: {traceback.format_exc()}")

                                        if fetched_email_objects:
                                            fetched_email_objects.sort(key=lambda em: em.get('receivedDateTime', ''))
                                            try:
                                                st.session_state.thread_structure = graph_client.build_thread_structure(fetched_email_objects)
                                            except Exception as e_build_struct:
                                                print(f"Could not build thread structure: {e_build_struct}")
                                                st.session_state.thread_structure = fetched_email_objects 
                                        
                                        if fetched_email_objects:
                                            st.write(f"Processing {len(fetched_email_objects)} fetched email(s)...")

                                            # --- Pre-flight check for AI Service and critical API keys ---
                                            ai_service = st.session_state.ai_service_instance
                                            gemini_key_present = bool(st.session_state.get("gemini_api_key"))
                                            # Dwellworks endpoint check is implicit in ai_service initialization
                                            # dwellworks_key_present = bool(st.session_state.get("dwellworks_api_key")) # REMOVED

                                            if not ai_service or not gemini_key_present: # Simplified check
                                                missing_keys_msgs = []
                                                if not ai_service:
                                                    missing_keys_msgs.append("AI Service could not be initialized (check Dwellworks Endpoint and Gemini Key).")
                                                if not gemini_key_present:
                                                    missing_keys_msgs.append("Gemini API Key is missing.")
                                                # if not dwellworks_key_present: # REMOVED
                                                #     missing_keys_msgs.append("Dwellworks API Key is missing.") 
                                                
                                                error_message_critical = "Critical API settings are missing. Please configure them in 'API Settings' in the sidebar: " + ", ".join(missing_keys_msgs)
                                                st.error(error_message_critical) 
                                                error_placeholder_api.empty() # Clear any previous transient messages
                                                progress_bar_api.empty()
                                                # Do not proceed with the loop
                                            else:
                                                # --- Proceed with processing loop if all checks pass ---
                                                for idx_email_proc, email_to_proc in enumerate(fetched_email_objects):
                                                    progress_bar_api.progress( (idx_email_proc + 1) / len(fetched_email_objects), text=f"Processing email {idx_email_proc + 1}/{len(fetched_email_objects)}...")
                                                    
                                                    # ai_service is already confirmed to be not None here
                                                    if gemini_key_present: # genai.configure should only be called if key is present
                                                        genai.configure(api_key=st.session_state.gemini_api_key)
                                                    # No need for an else here to print, as the critical check above handles it.

                                                    result_item_api = process_individual_email(email_to_proc, ai_service, previous_summary_api)
                                                    
                                                    if result_item_api:
                                                        result_item_api["email_index"] = idx_email_proc + 1 
                                                        results_api.append(result_item_api)
                                                        if isinstance(result_item_api.get("ai_output"), dict) and "Summary" in result_item_api["ai_output"]:
                                                            previous_summary_api = result_item_api["ai_output"]["Summary"]
                                                        if isinstance(result_item_api.get("ai_output"), dict) and result_item_api["ai_output"].get("error"):
                                                            if "rate limit" in result_item_api["ai_output"].get("error", "").lower():
                                                                has_rate_limit_error_api = True
                                        else:
                                            st.warning("No email objects could be successfully fetched for processing from this thread selection.")
                                    
                                    progress_bar_api.empty()
                                    if results_api:
                                        ensure_results_tab_works(results_api)
                                        if 'individual_results' in st.session_state and st.session_state.individual_results:
                                            if has_rate_limit_error_api:
                                                error_placeholder_api.error("⚠️ Some API emails had processing issues due to rate limits.")
                                            else:
                                                st.success(f"Successfully analyzed {len(results_api)} emails from Graph API!")
                                        else: 
                                            error_placeholder_api.error("Error storing Graph API results after processing.") # More specific error
                                    else: 
                                        # This warning is now covered by the one inside the 'else' for 'if not message_ids_to_fetch:'
                                        # or 'if fetched_email_objects:' else block. Avoid duplicate warnings.
                                        # st.warning("No emails processed from Graph API.") 
                                        if not message_ids_to_fetch: # Only show if initial list was empty
                                             pass # Warning already shown
                                        elif not fetched_email_objects: # Only show if fetching failed
                                             pass # Warning already shown
                                        else: # If we had IDs, fetched objects, but results_api is still empty
                                            st.warning("Emails were fetched but no results were generated from processing.")

                                        if 'progress_bar_api' in locals() and progress_bar_api is not None:
                                            progress_bar_api.empty()
                        except Exception as e_main_proc:
                            st.error(f"An error occurred while processing the Graph API thread: {str(e_main_proc)}")
                            traceback.print_exc()
                            if 'progress_bar_api' in locals() and progress_bar_api is not None:
                                progress_bar_api.empty()
        else: 
             st.info("No Outlook threads fetched yet. Use the filters and 'Fetch Email Threads' button above.")

    elif email_source == "Upload Email Excel":
        st.subheader("Upload Email Content from Excel")
        # ... existing code ...

# Add the calculate_similarity function before display_evaluation_metrics
def calculate_similarity(text1, text2):
    """Calculate similarity between two text strings using cosine similarity"""
    if not text1 or not text2:
        return 0.0
    
    # Tokenize and create sets of words
    import re
    words1 = set(re.findall(r'\b\w+\b', text1.lower()))
    words2 = set(re.findall(r'\b\w+\b', text2.lower()))
    
    # Calculate Jaccard similarity (intersection over union)
    intersection = len(words1.intersection(words2))
    union = len(words1.union(words2))
    
    if union == 0:
        return 0.0
    
    return intersection / union

def display_formatted_json(data, title=None):
    """
    Display JSON data with proper formatting and styling
    
    Args:
        data: The JSON data to display
        title: Optional title to show above the JSON
    """
    if title:
        st.markdown(f"### {title}")
    
    # Convert to JSON string if necessary
    if not isinstance(data, str):
        try:
            json_str = json.dumps(data, indent=2)
        except Exception as e:
            st.error(f"Error formatting JSON: {str(e)}")
            return
    else:
        json_str = data
    
    # Display using st.json for clean formatting
    st.json(data)

def highlight_status(val):
    """Helper function to highlight Pass/Fail/Partial Pass status with colors"""
    if val == 'Pass':
        return 'background-color: #008000; color: white'
    elif val == 'Fail':
        return 'background-color: #FF0000; color: white'
    elif val == 'Partial Pass':
        return 'background-color: #FFA500; color: black'
    elif val == 'Info':
        return 'background-color: #A9A9A9; color: white'
    else:
        return ''

def display_evaluation_metrics(result):
    """Display evaluation metrics in a structured way"""
    try:
        # Extract data from result
        ai_output = result.get('ai_output', {})
        groundtruth = result.get('groundtruth', {})
        metrics = result.get('metrics', [])
        
        if not ai_output:
            st.error("No AI output available to evaluate")
            return
            
        # Display email summary for context
        with st.expander("📝 Email Summary", expanded=True):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### 📊 AI Output")
                ai_summary = ai_output.get('Summary', 'No summary available')
                st.info(ai_summary)
                
            with col2:
                st.markdown("### 🎯 Ground Truth")
                gt_summary = groundtruth.get('Summary', 'No ground truth available')
                st.success(gt_summary)
                
            # Show evaluation error if present
            if 'evaluation_error' in result and result['evaluation_error']:
                st.error(f"Evaluation Error: {result['evaluation_error']}")
        
        # If no metrics but evaluation error exists, we should generate fallback metrics
        if not metrics and result.get('evaluation_error') and ai_output and groundtruth:
            print("Generating fallback metrics from evaluation error")
            try:
                from evaluation_engine import generate_fallback_metrics
                metrics = generate_fallback_metrics(ai_output, groundtruth)
                print(f"Generated {len(metrics)} fallback metrics")
            except Exception as e:
                print(f"Error generating fallback metrics: {str(e)}")
        
        # Group metrics by type
        summary_metrics = []
        sentiment_metrics = []
        feature_metrics = []
        event_metrics = []
        
        for metric in metrics:
            metric_name = metric.get('Metric', metric.get('Field', '')).lower()
            
            if 'summary' in metric_name:
                summary_metrics.append(metric)
            elif 'sentiment' in metric_name:
                sentiment_metrics.append(metric)
            elif any(x in metric_name for x in ['feature', 'category']):
                feature_metrics.append(metric)
            elif 'event' in metric_name:
                event_metrics.append(metric)
        
        # Helper function to render metrics in the desired table format
        def render_metric_display_table(metric_dict, metric_name_for_title):
            import pandas as pd
            import streamlit as st
            st.subheader(metric_name_for_title.replace("_", " ").title())

            table_data = []
            ai_value = metric_dict.get("AI Value", "N/A")
            gt_value = metric_dict.get("Ground Truth", "N/A")

            # --- START EVENT HANDLING ---
            # If it's event data and the value is a list or dict, convert to JSON string for display
            if "event" in metric_name_for_title.lower():
                if isinstance(ai_value, (list, dict)):
                    try:
                        ai_value = json.dumps(ai_value, indent=2)
                    except Exception:
                        ai_value = str(ai_value) # Fallback to string conversion
                if isinstance(gt_value, (list, dict)):
                    try:
                        gt_value = json.dumps(gt_value, indent=2)
                    except Exception:
                        gt_value = str(gt_value) # Fallback to string conversion
            # --- END EVENT HANDLING ---
            
            # Always add AI Value and Ground Truth (potentially formatted)
            table_data.append({"Metric": "AI Value", "Value": ai_value})
            table_data.append({"Metric": "Ground Truth", "Value": gt_value})

            # Add Similarity Percentage ONLY for Summary metric
            if metric_name_for_title.lower() == "summary":
                similarity_val = metric_dict.get("Similarity Percentage", metric_dict.get("Similarity")) # Check both keys
                # Basic formatting to ensure it looks like a percentage string
                if isinstance(similarity_val, (int, float)):
                     similarity_val = f"{similarity_val:.0%}"
                elif isinstance(similarity_val, str) and "%" not in similarity_val:
                     try:
                         similarity_float = float(similarity_val)
                         similarity_val = f"{similarity_float:.0%}"
                     except ValueError:
                         pass # Keep string as is if conversion fails
                table_data.append({"Metric": "Similarity Percentage", "Value": similarity_val if similarity_val is not None else "N/A"})

            # Add Pass/Fail Status
            table_data.append({"Metric": "Pass/Fail", "Value": metric_dict.get("Status", "N/A")})
            
            # Add Ground Truth Explanation (Mandatory - should not be N/A from LLM)
            table_data.append({"Metric": "Ground Truth Explanation", "Value": metric_dict.get("Ground Truth Explanation", "Explanation missing from LLM")})
            
            # Add Pass/Fail or % Explanation
            explanation_key = "% Explanation" if metric_name_for_title.lower() == "summary" else "Pass/Fail Explanation"
            explanation_value = metric_dict.get(explanation_key, metric_dict.get("Pass/Fail Explanation", metric_dict.get("Comparison Explanation", "Explanation missing from LLM")))
            table_data.append({"Metric": explanation_key, "Value": explanation_value})
            
            df = pd.DataFrame(table_data)
            print(f"DEBUG render_metric_display_table: Constructed DataFrame for '{metric_name_for_title}':\n{df.to_string()}")
            
            st.table(df.set_index("Metric"))

        # 1. SENTIMENT ANALYSIS EVALUATION - in an expander
        with st.expander("📊 SENTIMENT ANALYSIS EVALUATION", expanded=True):
            sentiment_actual_metric = next((m for m in sentiment_metrics if m.get('Metric', m.get('Field', '')).strip().lower() == 'sentiment analysis'), None)
            overall_sentiment_metric = next((m for m in sentiment_metrics if m.get('Metric', m.get('Field', '')).strip().lower() == 'overall_sentiment_analysis'), None)
            
            if sentiment_actual_metric:
                render_metric_display_table(sentiment_actual_metric, "Sentiment Analysis")
            else:
                st.info("No Sentiment Analysis (red/green) metrics available.")
                
            if overall_sentiment_metric:
                render_metric_display_table(overall_sentiment_metric, "Overall Sentiment Analysis")
            else:
                st.info("No Overall Sentiment Analysis (positive/negative/neutral) metrics available.")
            
            if not sentiment_actual_metric and not overall_sentiment_metric:
                st.info("No sentiment analysis metrics available for this email.")

        # 2. FEATURE & CATEGORY EVALUATION - in an expander
        with st.expander("🔍 FEATURE & CATEGORY EVALUATION", expanded=True):
            feature_metric = next((m for m in feature_metrics if m.get('Metric', m.get('Field', '')).strip().lower() == 'feature'), None)
            category_metric = next((m for m in feature_metrics if m.get('Metric', m.get('Field', '')).strip().lower() == 'category'), None)

            if feature_metric:
                render_metric_display_table(feature_metric, "Feature")
            else:
                st.info("No Feature metrics available.")

            if category_metric:
                render_metric_display_table(category_metric, "Category")
            else:
                st.info("No Category metrics available.")

            if not feature_metric and not category_metric:
                st.info("No feature or category metrics available for this email.")
            
            # Add feature classification matrix as reference - this can stay as is or be removed if too cluttered
            # st.markdown("### Feature Classification Matrix Reference")
            # feature_matrix = pd.DataFrame([...]) # Definition from existing code
            # st.dataframe(feature_matrix)

        # 3. EVENT DETECTION EVALUATION - in an expander
        with st.expander("🗓️ EVENT DETECTION EVALUATION", expanded=True):
            # Try to find an overall event metric first (e.g., Event Match, Events_count, or just Events)
            event_summary_metric = next((m for m in event_metrics if m.get('Metric', m.get('Field', '')).strip().lower() in ['event match', 'events_count', 'events']), None)
            
            if event_summary_metric:
                # Adjust AI/Ground Truth value if it's a count for display
                # The user example shows count for AI Value and Ground Truth for events
                ai_val = event_summary_metric.get("AI Value", "N/A")
                gt_val = event_summary_metric.get("Ground Truth", "N/A")
                
                # If values look like counts (e.g., "1 events detected"), extract the number
                if isinstance(ai_val, str) and "events detected" in ai_val.lower():
                    ai_val_display = ai_val.lower().split(" ")[0]
                else:
                    ai_val_display = ai_val

                if isinstance(gt_val, str) and "events" in gt_val.lower(): # Ground truth might be like "1 events" or "1 events in groundtruth"
                    gt_val_display = gt_val.lower().split(" ")[0]
                else:
                    gt_val_display = gt_val
                
                # Create a copy to modify for display without affecting original dict
                display_metric_dict = event_summary_metric.copy()
                display_metric_dict["AI Value"] = ai_val_display
                display_metric_dict["Ground Truth"] = gt_val_display
                
                render_metric_display_table(display_metric_dict, "Event Detection")
                
                # Display individual event field details if they exist and are separate metrics
                # This part might need adjustment based on how event_metrics_list is structured
                # For now, focusing on the main table as per user's primary example

            else:
                st.info("No event detection summary metrics available for this email.")

        # 4. SUMMARY EVALUATION - in an expander
        with st.expander("📝 SUMMARY EVALUATION", expanded=True):
            summary_metric = next((m for m in summary_metrics if m.get('Metric', m.get('Field', '')).strip().lower() == 'summary'), None)
            
            if summary_metric:
                render_metric_display_table(summary_metric, "Summary")
            else:
                st.info("No summary evaluation metrics available for this email.")
        
    except Exception as e:
        st.error(f"Error displaying evaluation metrics: {str(e)}")
        logging.exception("Error in display_evaluation_metrics")

def evaluate_results(ai_output, groundtruth=None):
    """
    Evaluate AI output against groundtruth to generate metrics
    
    Args:
        ai_output: Dictionary containing AI-generated output
        groundtruth: Dictionary containing groundtruth data
        
    Returns:
        List of metrics dictionaries with evaluation results
    """
    print("=== Debug: Starting evaluate_results ===")
    
    if not ai_output:
        return []
    
    # Initialize metrics list
    metrics = []
    
    # If no groundtruth is provided, we can only do basic validation
    if not groundtruth:
        # Basic validation metrics
        metrics.append({
            "Field": "validation",
            "Category": "Format",
            "AI Value": "Complete",
            "Ground Truth": "Unknown",
            "Pass/Fail": "Pass",
            "Pass/Fail Explanation": "AI output was successfully generated with all required fields"
        })
        return metrics
    
    # 1. Sentiment Analysis
    ai_sentiment = ai_output.get("Sentiment analysis", "")
    gt_sentiment = groundtruth.get("Sentiment analysis", "")
    
    sentiment_pass = ai_sentiment == gt_sentiment
    
    # Build a more detailed explanation for sentiment analysis
    if sentiment_pass:
        sentiment_explanation = (
            f"The AI correctly identified the sentiment as '{ai_sentiment}'. "
            f"This matches the groundtruth sentiment. "
        )
        if ai_sentiment == "green":
            sentiment_explanation += (
                "The email has a positive tone, containing elements like appreciation, "
                "gratitude, good news, or friendly language."
            )
        else:  # red
            sentiment_explanation += (
                "The email has a negative tone, containing elements like complaints, "
                "issues, apologies, delays, or expressions of frustration."
            )
    else:
        sentiment_explanation = (
            f"The AI identified the sentiment as '{ai_sentiment}' but the groundtruth indicates '{gt_sentiment}'. "
        )
        if gt_sentiment == "green":
            sentiment_explanation += (
                "The email actually has a positive tone, containing elements like appreciation, "
                "gratitude, good news, or friendly language that the AI failed to recognize."
            )
        else:  # red
            sentiment_explanation += (
                "The email actually has a negative tone, containing elements like complaints, "
                "issues, apologies, delays, or expressions of frustration that the AI failed to recognize."
            )
    
    metrics.append({
        "Field": "Sentiment analysis",
        "Category": "Basic Analysis",
        "AI Value": ai_sentiment,
        "Ground Truth": gt_sentiment,
        "Pass/Fail": "Pass" if sentiment_pass else "Fail",
        "Pass/Fail Explanation": sentiment_explanation
    })
    
    # 2. Overall Sentiment Analysis
    ai_overall = ai_output.get("overall_sentiment_analysis", "")
    gt_overall = groundtruth.get("overall_sentiment_analysis", "")
    
    overall_pass = ai_overall == gt_overall
    
    # Build a more detailed explanation for overall sentiment
    if overall_pass:
        overall_explanation = (
            f"The AI correctly identified the overall sentiment as '{ai_overall}'. "
            f"This matches the groundtruth assessment. "
        )
        if ai_overall == "positive":
            overall_explanation += (
                "The email has a generally positive tone, expressing friendly, appreciative, "
                "grateful, or excited sentiments."
            )
        elif ai_overall == "negative":
            overall_explanation += (
                "The email has a generally negative tone, expressing frustration, annoyance, "
                "disappointment, or containing apologetic language."
            )
        else:  # neutral
            overall_explanation += (
                "The email has a neutral tone, primarily containing informational content "
                "without strong positive or negative emotional elements."
            )
    else:
        overall_explanation = (
            f"The AI identified the overall sentiment as '{ai_overall}' but the groundtruth indicates '{gt_overall}'. "
        )
        if gt_overall == "positive":
            overall_explanation += (
                "The email actually has a positive tone that the AI missed, expressing friendly, "
                "appreciative, grateful, or excited sentiments."
            )
        elif gt_overall == "negative":
            overall_explanation += (
                "The email actually has a negative tone that the AI missed, expressing frustration, "
                "annoyance, disappointment, or containing apologetic language."
            )
        else:  # neutral
            overall_explanation += (
                "The email actually has a neutral tone that the AI missed, primarily containing "
                "informational content without strong positive or negative emotional elements."
            )
    
    metrics.append({
        "Field": "overall_sentiment_analysis",
        "Category": "Basic Analysis",
        "AI Value": ai_overall,
        "Ground Truth": gt_overall,
        "Pass/Fail": "Pass" if overall_pass else "Fail",
        "Pass/Fail Explanation": overall_explanation
    })
    
    # 3. Feature & Category
    ai_feature = ai_output.get("feature", "")
    gt_feature = groundtruth.get("feature", "")
    
    ai_category = ai_output.get("category", "")
    gt_category = groundtruth.get("category", "")
    
    feature_pass = ai_feature == gt_feature
    category_pass = ai_category == gt_category
    
    # Build a more detailed explanation for feature identification
    valid_features = [
        "EMAIL -- DSC First Contact with EE Completed", 
        "EMAIL -- EE First Contact with DSC",
        "EMAIL -- Phone Consultation Scheduled", 
        "EMAIL -- Phone Consultation Completed",
        "no feature"
    ]
    
    feature_descriptions = {
        "EMAIL -- DSC First Contact with EE Completed": "This is a first email sent by DSC to EE",
        "EMAIL -- EE First Contact with DSC": "This is a first email received by DSC from EE",
        "EMAIL -- Phone Consultation Scheduled": "The email mentions a future phone consultation",
        "EMAIL -- Phone Consultation Completed": "The email indicates a phone consultation was completed",
        "no feature": "None of the specific features apply to this email"
    }
    
    if feature_pass:
        feature_explanation = (
            f"The AI correctly identified the feature as '{ai_feature}'. "
            f"This matches the groundtruth feature. "
        )
        if ai_feature in feature_descriptions:
            feature_explanation += feature_descriptions[ai_feature]
    else:
        feature_explanation = (
            f"The AI identified the feature as '{ai_feature}' but the groundtruth indicates '{gt_feature}'. "
        )
        if gt_feature in feature_descriptions:
            feature_explanation += (
                f"The email should be classified as '{gt_feature}' because: {feature_descriptions[gt_feature]}"
            )
    
    metrics.append({
        "Field": "feature",
        "Category": "Classification",
        "AI Value": ai_feature,
        "Ground Truth": gt_feature,
        "Pass/Fail": "Pass" if feature_pass else "Fail",
        "Pass/Fail Explanation": feature_explanation
    })
    
    # Build a more detailed explanation for category
    category_rules = {
        "Initial Service Milestones": "This category applies to any of the specific email features",
        "no category": "This applies when the feature is 'no feature'"
    }
    
    if category_pass:
        category_explanation = (
            f"The AI correctly identified the category as '{ai_category}'. "
            f"This matches the groundtruth category. "
        )
        if ai_category in category_rules:
            category_explanation += category_rules[ai_category]
    else:
        category_explanation = (
            f"The AI identified the category as '{ai_category}' but the groundtruth indicates '{gt_category}'. "
        )
        if gt_category in category_rules:
            category_explanation += (
                f"The correct category should be '{gt_category}' because: {category_rules[gt_category]}"
            )
        
        # Add explanation about the relationship between feature and category
        if gt_feature != "no feature" and gt_category != ai_category:
            category_explanation += (
                f" Since the feature is '{gt_feature}', the category should be 'Initial Service Milestones'."
            )
        elif gt_feature == "no feature" and gt_category != ai_category:
            category_explanation += (
                " Since the feature is 'no feature', the category should be 'no category'."
            )
    
    metrics.append({
        "Field": "category",
        "Category": "Classification",
        "AI Value": ai_category,
        "Ground Truth": gt_category,
        "Pass/Fail": "Pass" if category_pass else "Fail",
        "Pass/Fail Explanation": category_explanation
    })
    
    # 4. Summary Evaluation
    ai_summary = ai_output.get("Summary", "")
    gt_summary = groundtruth.get("Summary", "")
    
    # Calculate similarity score
    summary_similarity = calculate_similarity(ai_summary, gt_summary)
    summary_percentage = int(summary_similarity * 100)
    
    # Get the explanation for why groundtruth summary is correct (if available)
    gt_summary_explanation = groundtruth.get("Summary_explanation", "")
    
    # Create a detailed explanation of the summary evaluation
    if summary_percentage >= 70:
        summary_explanation = (
            f"The AI summary captures the essential information with {summary_percentage}% similarity to the groundtruth. "
            "It correctly uses indirect speech style and includes the key points from the email. "
            f"The AI summary received a PASS rating because it exceeds the 70% similarity threshold."
        )
    elif summary_percentage >= 50:
        missing_percent = 70 - summary_percentage
        summary_explanation = (
            f"The AI summary partially captures the information with {summary_percentage}% similarity to the groundtruth. "
            f"It needs {missing_percent}% more similarity for a full pass. "
            "The summary may be missing some key details or not fully using indirect speech style throughout. "
            "It received a PARTIAL PASS rating."
        )
    else:
        missing_percent = 70 - summary_percentage
        summary_explanation = (
            f"The AI summary has LOW similarity ({summary_percentage}%) with the groundtruth. "
            f"It needs {missing_percent}% more similarity for a passing grade. "
            "The summary may have significant omissions, use direct speech instead of indirect speech, "
            "or contain inaccuracies compared to the email content. "
            "It received a FAIL rating."
        )
    
    # Add more explanation about indirect speech if not already included
    if "indirect speech" not in summary_explanation.lower():
        if "Direct speech" in ai_summary or "I " in ai_summary or "We " in ai_summary:
            summary_explanation += " The AI summary appears to use direct speech in some places, which should be avoided."
    
    # Generate an explanation about the groundtruth if one isn't provided
    if not gt_summary_explanation:
        gt_summary_explanation = (
            "The groundtruth summary was generated following best practices for email summarization: "
            "1) It uses proper indirect speech throughout, avoiding first and second person pronouns. "
            "2) It captures all essential information from the email. "
            "3) It maintains the same level of formality as the original email. "
            "4) It uses neutral reporting phrases like 'the sender mentioned' and 'they stated'."
        )
    
    # Determine pass/fail based on similarity threshold
    summary_status = "Pass" if summary_percentage >= 70 else "Partial Pass" if summary_percentage >= 50 else "Fail"
    
    metrics.append({
        "Field": "Summary",
        "Category": "Content Analysis",
        "AI Value": ai_summary,
        "Ground Truth": gt_summary,
        "Pass/Fail": summary_status,
        "Pass/Fail Explanation": summary_explanation,
        "Similarity": f"{summary_percentage}%",
        "GT Explanation": gt_summary_explanation
    })
    
    # 5. Events Detection
    ai_events = ai_output.get("Events", [])
    gt_events = groundtruth.get("Events", [])
    
    # Count matching events
    matching_events = 0
    partial_matches = 0
    
    # Detailed analysis of event matches
    event_details = []
    
    # Check if events match
    events_match = True
    if len(ai_events) != len(gt_events):
        events_match = False
        event_details.append(f"Count mismatch: AI detected {len(ai_events)} events while groundtruth has {len(gt_events)}")
    
    # If both are empty lists, it's a match
    if len(ai_events) == 0 and len(gt_events) == 0:
        events_match = True
        event_details.append("Both AI and groundtruth correctly identified no events in the email")
    else:
        # Detailed comparison of each event
        for i, gt_event in enumerate(gt_events):
            if i >= len(ai_events):
                event_details.append(f"Missing event in AI output: {gt_event.get('Event name', 'Unknown event')}")
                continue
                
            ai_event = ai_events[i]
            match_count = 0
            total_fields = 0
            
            # Compare individual fields for this event
            field_analysis = []
            for field in ["Event name", "Date", "Time", "Property Type", "Agent Name", "Location"]:
                ai_value = ai_event.get(field, "null")
                gt_value = gt_event.get(field, "null")
                
                # Skip null-null matches for non-essential fields
                if ai_value == "null" and gt_value == "null" and field not in ["Event name", "Date", "Time"]:
                    continue
                    
                total_fields += 1
                if ai_value == gt_value:
                    match_count += 1
                else:
                    field_analysis.append(f"{field}: AI='{ai_value}', GT='{gt_value}'")
            
            # Calculate match percentage
            match_percentage = (match_count / total_fields * 100) if total_fields > 0 else 0
            
            if match_percentage == 100:
                matching_events += 1
                event_details.append(f"Event {i+1} ({gt_event.get('Event name', 'Unknown')}): Perfect match (100%)")
            elif match_percentage >= 70:
                partial_matches += 1
                event_details.append(f"Event {i+1} ({gt_event.get('Event name', 'Unknown')}): Partial match ({match_percentage:.0f}%). Differences: {', '.join(field_analysis)}")
            else:
                event_details.append(f"Event {i+1} ({gt_event.get('Event name', 'Unknown')}): Low match ({match_percentage:.0f}%). Differences: {', '.join(field_analysis)}")
        
        # Check if AI detected extra events not in groundtruth
        if len(ai_events) > len(gt_events):
            for i in range(len(gt_events), len(ai_events)):
                ai_event = ai_events[i]
                event_details.append(f"Extra event in AI output: {ai_event.get('Event name', 'Unknown event')}")
    
    # Determine overall event detection status
    if events_match:
        events_status = "Pass"
    elif (matching_events + partial_matches) == len(gt_events) and len(ai_events) == len(gt_events):
        events_status = "Partial Pass"
    else:
        events_status = "Fail"
    
    # Create a comprehensive explanation
    if events_status == "Pass":
        if len(ai_events) == 0:
            events_explanation = "The AI correctly identified that there are no events mentioned in the email."
        else:
            events_explanation = f"The AI correctly identified all {len(ai_events)} events with accurate details for each event."
    elif events_status == "Partial Pass":
        events_explanation = (
            f"The AI correctly identified the number of events ({len(gt_events)}), "
            f"with {matching_events} perfect matches and {partial_matches} partial matches. "
            "Some event details may have minor differences."
        )
    else:
        if len(ai_events) == 0 and len(gt_events) > 0:
            events_explanation = f"The AI failed to detect any events, while the groundtruth contains {len(gt_events)} events."
        elif len(ai_events) > 0 and len(gt_events) == 0:
            events_explanation = f"The AI incorrectly detected {len(ai_events)} events, while the email doesn't actually contain any events."
        elif len(ai_events) != len(gt_events):
            events_explanation = f"The AI detected {len(ai_events)} events while the groundtruth has {len(gt_events)} events."
        else:
            events_explanation = "The AI detected the correct number of events, but with significant differences in the details."
    
    # Add detailed analysis to the explanation
    if event_details:
        events_explanation += "\n\nDetailed event analysis:\n- " + "\n- ".join(event_details)
    
    # Create events metric
    metrics.append({
        "Field": "Events",
        "Category": "Event Detection",
        "AI Value": json.dumps(ai_events, indent=2),
        "Ground Truth": json.dumps(gt_events, indent=2),
        "Pass/Fail": events_status,
        "Pass/Fail Explanation": events_explanation
    })
    
    return metrics

# ------------------------------------------------------------
# Tab 2: Detailed Evaluation
# ------------------------------------------------------------

with tab2:
    st.header("Detailed Evaluation Metrics")
    
    # Check if we have results to display
    has_results = (
        'individual_results' in st.session_state 
        and isinstance(st.session_state.individual_results, list)
        and len(st.session_state.individual_results) > 0
    )
    
    if has_results:
        individual_results = st.session_state.individual_results
        
        # --- ADD DOWNLOAD BUTTON HERE ---
        excel_data = convert_metrics_to_excel(individual_results)
        # Construct a dynamic filename
        thread_subject = get_clean_subject(individual_results[0].get('email', {}).get('subject', 'Thread'))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"evaluation_metrics_{thread_subject[:30].replace(' ', '_')}_{timestamp}.xlsx"
        
        st.download_button(
            label="📥 Download Metrics as Excel",
            data=excel_data,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("---") # Add a separator
        # --- END DOWNLOAD BUTTON ---
        
        # --- START DASHBOARD DISPLAY ---
        st.subheader("Thread Performance Dashboard")
        total_mails = len(individual_results)
        total_fields_validated = 0
        total_pass = 0

        for result in individual_results:
            metrics_for_email = result.get('metrics', [])
            for metric_item in metrics_for_email:
                # Consider each top-level metric as one field initially
                total_fields_validated += 1
                if metric_item.get("Status") == "Pass":
                    total_pass += 1
                
                # If it's an event metric, and AI/GT values are lists (of event dicts), count sub-fields
                # This is a simplified approach; true sub-field count might need more complex parsing if events are not structured
                if "event" in metric_item.get("Metric", metric_item.get("Field", "")).lower():
                    ai_event_val = metric_item.get("AI Value")
                    gt_event_val = metric_item.get("Ground Truth")
                    # If AI Value was a list of events, it's now a JSON string. Try to parse back.
                    # This is a basic check. A more robust way would be to count fields from the original event dicts if available.
                    try:
                        if isinstance(ai_event_val, str):
                            parsed_ai_events = json.loads(ai_event_val)
                            if isinstance(parsed_ai_events, list) and len(parsed_ai_events) > 0:
                                # Assuming each event dict has ~5-6 relevant sub-fields
                                total_fields_validated += (len(parsed_ai_events[0]) -1) * len(parsed_ai_events) # -1 for the main event, then add per event
                    except: # json.JSONDecodeError or other issues
                        pass # Keep initial count if parsing fails
        
        accuracy_percentage = (total_pass / total_fields_validated * 100) if total_fields_validated > 0 else 0

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Mails Processed", total_mails)
        with col2:
            st.metric("Total Fields Validated", total_fields_validated)
        with col3:
            st.metric("Total Fields Passed", total_pass)
        with col4:
            st.metric("Overall Accuracy", f"{accuracy_percentage:.2f}%")
        st.markdown("---")
        # --- END DASHBOARD DISPLAY ---
        
        # --- CREATE TABS INCLUDING OVERALL REVIEW ---
        email_tab_titles = [
            f"Email {idx+1}: {get_clean_subject(result.get('email', {}).get('subject', 'No Subject'))[:30]}..." 
            for idx, result in enumerate(individual_results)
        ]
        # Add the Overall Review tab title AT THE BEGINNING
        all_tab_titles = ["🔍 Overall Thread Review"] + email_tab_titles
        
        # Ensure tab_titles are unique to prevent TypeError
        unique_all_tab_titles = list(dict.fromkeys(all_tab_titles))
        
        # Create the tabs ONCE
        created_tabs = st.tabs(unique_all_tab_titles)
        
        # --- Display Overall Thread Review Tab (Now the first tab) ---
        with created_tabs[0]: 
            st.header("Overall Thread Review")
            with st.spinner("Generating overall thread review..."):
                overall_review_text = generate_overall_thread_review(individual_results)
                st.markdown(overall_review_text)

        # --- Display Individual Email Evaluations (Starting from the second tab) ---
        for idx, email_tab in enumerate(created_tabs[1:]): # Loop through email tabs
            with email_tab:
                result = individual_results[idx] # Correctly index individual_results
                try:
                    st.subheader(f"Email {idx+1} Detailed Evaluation")
                    sender_name = result.get('email', {}).get('from', {}).get('emailAddress', {}).get('name', 'Unknown')
                    sender_email = result.get('email', {}).get('from', {}).get('emailAddress', {}).get('address', 'Unknown')
                    subject = result.get('email', {}).get('subject', 'No Subject')
                    st.markdown(f"**From:** {sender_name} ({sender_email})")
                    st.markdown(f"**Subject:** {subject}")
                    
                    metrics = result.get("metrics", [])
                    
                    total_metrics_calc = len([m for m in metrics if m.get("Status") != "Info" and m.get("Status") is not None])
                    pass_count_calc = sum(1 for m in metrics if m.get("Status") == "Pass")
                    partial_count_calc = sum(1 for m in metrics if m.get("Status") == "Partial Pass")

                    if total_metrics_calc > 0:
                        pass_percentage_calc = (pass_count_calc + (partial_count_calc * 0.5)) / total_metrics_calc * 100
                    else:
                        pass_percentage_calc = 0.0

                    if total_metrics_calc == 0:
                        st.info("ℹ️ Overall Evaluation for this Email: No metrics to score")
                    elif pass_percentage_calc >= 80:
                        st.success("✅ Overall Evaluation for this Email: PASSED")
                    elif pass_percentage_calc >= 50:
                        st.warning("⚠️ Overall Evaluation for this Email: PARTIALLY PASSED")
                    else:
                        st.error("❌ Overall Evaluation for this Email: FAILED")
                    
                    sentiment_metrics_list = [m for m in metrics if m.get('Metric', m.get('Field', '')).strip().lower() in ['sentiment analysis', 'overall_sentiment_analysis']]
                    feature_metrics_list = [m for m in metrics if m.get('Metric', m.get('Field', '')).strip().lower() in ['feature', 'category']]
                    summary_metrics_list = [m for m in metrics if m.get('Metric', m.get('Field', '')).strip().lower() in ['summary']]
                    event_metrics_list = [m for m in metrics if 'event' in m.get('Metric', m.get('Field', '')).strip().lower()]

                    with st.expander("📊 SENTIMENT ANALYSIS EVALUATION", expanded=True):
                        sentiment_actual_metric = next((m for m in sentiment_metrics_list if m.get('Metric', m.get('Field', '')).strip().lower() == 'sentiment analysis'), None)
                        overall_sentiment_metric = next((m for m in sentiment_metrics_list if m.get('Metric', m.get('Field', '')).strip().lower() == 'overall_sentiment_analysis'), None)
                        if sentiment_actual_metric: render_metric_display_table(sentiment_actual_metric, "Sentiment Analysis") 
                        else: st.info("No Sentiment Analysis (red/green) metrics available.")
                        if overall_sentiment_metric: render_metric_display_table(overall_sentiment_metric, "Overall Sentiment Analysis")
                        else: st.info("No Overall Sentiment Analysis (positive/negative/neutral) metrics available.")
                        if not sentiment_actual_metric and not overall_sentiment_metric: st.info("No sentiment analysis metrics available for this email.")

                    with st.expander("🔍 FEATURE & CATEGORY EVALUATION", expanded=True):
                        feature_metric = next((m for m in feature_metrics_list if m.get('Metric', m.get('Field', '')).strip().lower() == 'feature'), None)
                        category_metric = next((m for m in feature_metrics_list if m.get('Metric', m.get('Field', '')).strip().lower() == 'category'), None)
                        if feature_metric: render_metric_display_table(feature_metric, "Feature")
                        else: st.info("No Feature metrics available.")
                        if category_metric: render_metric_display_table(category_metric, "Category")
                        else: st.info("No Category metrics available.")
                        if not feature_metric and not category_metric: st.info("No feature or category metrics available for this email.")

                    with st.expander("🗓️ EVENT DETECTION EVALUATION", expanded=True):
                        event_summary_metric = next((m for m in event_metrics_list if m.get('Metric', m.get('Field', '')).strip().lower() in ['event match', 'events_count', 'events']), None)
                        if event_summary_metric:
                            ai_val = event_summary_metric.get("AI Value", "N/A")
                            gt_val = event_summary_metric.get("Ground Truth", "N/A")
                            if isinstance(ai_val, str) and "events detected" in ai_val.lower(): ai_val_display = ai_val.lower().split(" ")[0]
                            else: ai_val_display = ai_val
                            if isinstance(gt_val, str) and "events" in gt_val.lower(): gt_val_display = gt_val.lower().split(" ")[0]
                            else: gt_val_display = gt_val
                            display_metric_dict = event_summary_metric.copy()
                            display_metric_dict["AI Value"] = ai_val_display
                            display_metric_dict["Ground Truth"] = gt_val_display
                            render_metric_display_table(display_metric_dict, "Event Detection")
                        else: st.info("No event detection summary metrics available for this email.")

                    with st.expander("📝 SUMMARY EVALUATION", expanded=True):
                        summary_metric = next((m for m in summary_metrics_list if m.get('Metric', m.get('Field', '')).strip().lower() == 'summary'), None)
                        if summary_metric: render_metric_display_table(summary_metric, "Summary")
                        else: st.info("No summary evaluation metrics available for this email.")
                        
                    # --- Display Individual Email Review Points --- 
                    st.markdown("#### Individual Email Review Points")
                    review_points_found = False
                    for metric_item_for_review in metrics:
                        if "individual_email_review_points" in metric_item_for_review:
                            metric_name_for_review = metric_item_for_review.get("Metric", metric_item_for_review.get("Field", "Review"))
                            points_text = metric_item_for_review.get("individual_email_review_points", "No specific review points provided by LLM.")
                            if points_text and points_text != "No specific review points provided by LLM.":
                                st.markdown(f"**Review for {metric_name_for_review}:**")
                                st.markdown(points_text) # Assumes points_text is already formatted with newlines/bullets
                                review_points_found = True
                    if not review_points_found:
                        st.info("No detailed review points were generated for this email by the LLM.")
                    st.markdown("---")

                except Exception as e: # This is the corresponding except block
                    st.error(f"Error displaying email {idx + 1} evaluation: {str(e)}")
                    import traceback
                    traceback.print_exc()
        
    else:
        st.info("No evaluation results available. Please process emails first.")