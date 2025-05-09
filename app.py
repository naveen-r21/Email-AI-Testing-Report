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

# Helper function to create a styled header
def styled_header(text, color="#ff4b4b"):
    st.markdown(f"""
        <h2 style='color: {color};'>{text}</h2>
    """, unsafe_allow_html=True)

# Helper function to get clean subject by removing prefixes like Re:, Fw:, etc.
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
    """Converts the list of individual results into an Excel file buffer with multiple sheets."""
    import pandas as pd 
    from io import BytesIO
    import re
    import json

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
        
        # --- NEW Sheet 3: Email Data ---
        email_data_sheet = []
        for idx, result in enumerate(individual_results):
            email_subject = result.get('email', {}).get('subject', f'Email {idx+1} - No Subject')
            # Convert each complex data structure to JSON string for Excel
            input_data_json = json.dumps(result.get('input_data', {}), indent=2) if isinstance(result.get('input_data'), dict) else str(result.get('input_data', 'N/A'))
            output_data_json = json.dumps(result.get('ai_output', {}), indent=2) if isinstance(result.get('ai_output'), dict) else str(result.get('ai_output', 'N/A'))
            groundtruth_data_json = json.dumps(result.get('groundtruth', {}), indent=2) if isinstance(result.get('groundtruth'), dict) else str(result.get('groundtruth', 'N/A'))
            
            email_data_sheet.append({
                "Email Number": idx + 1,
                "Email Subject": get_clean_subject(email_subject),
                "Input Data": input_data_json,
                "Output Data": output_data_json,
                "Groundtruth Data": groundtruth_data_json
            })
        
        df_email_data = pd.DataFrame(email_data_sheet)
        df_email_data.to_excel(writer, index=False, sheet_name='Email Data')
        worksheet_email_data = writer.sheets['Email Data']
        text_wrap_format = writer.book.add_format({'text_wrap': True, 'valign': 'top'})
        
        # Set column widths for the Email Data sheet
        worksheet_email_data.set_column(0, 0, 15)  # Email Number
        worksheet_email_data.set_column(1, 1, 40)  # Email Subject
        worksheet_email_data.set_column(2, 2, 60, text_wrap_format)  # Input Data
        worksheet_email_data.set_column(3, 3, 60, text_wrap_format)  # Output Data
        worksheet_email_data.set_column(4, 4, 60, text_wrap_format)  # Groundtruth Data
            
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
    """Initialize all session state variables with default values"""
    defaults = {
        'individual_results': [],
        'has_results': False,
        'threads': [],
        'active_tab': None,
        'gemini_api_key': os.getenv("GEMINI_API_KEY", ""),
        'ms_graph_client_id': os.getenv("MS_GRAPH_CLIENT_ID", ""),
        'ms_graph_client_secret': os.getenv("MS_GRAPH_CLIENT_SECRET", ""),
        'ms_graph_tenant_id': os.getenv("MS_GRAPH_TENANT_ID", ""),
        'user_email_address': os.getenv("USER_EMAIL", ""),
        'email_source': "Upload emails as Excel",  # Changed default
        'current_page': 1,
        'items_per_page': 5,
        'selected_thread': None,
        'previous_summary': None,
        'evaluation_metrics': {},
        'error_message': None,
        'success_message': None
    }
    
    for key, default_value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value

    # --- DEBUG PRINT ADDED ---
    print(f"DEBUG: Inside initialize_session_state, GEMINI_API_KEY from env: '{os.getenv('GEMINI_API_KEY', '<NOT_SET_OR_EMPTY>')}'")
    print(f"DEBUG: Inside initialize_session_state, USER_EMAIL from env: '{os.getenv('USER_EMAIL', '<NOT_SET_OR_EMPTY>')}'")
    print(f"DEBUG: Session state gemini_api_key after init: '{st.session_state.get('gemini_api_key')}'")
    print(f"DEBUG: Session state user_email_address after init: '{st.session_state.get('user_email_address')}'")
    # --- END DEBUG PRINT ---

# Call initialization at startup
initialize_session_state()

# Helper function to display results dashboard and download button
def display_results_dashboard():
    """Display the metrics dashboard and download button after processing"""
    if not (
        'individual_results' in st.session_state 
        and isinstance(st.session_state.individual_results, list)
        and len(st.session_state.individual_results) > 0
    ):
        return
    
    individual_results = st.session_state.individual_results
    
    # --- ADD DOWNLOAD BUTTON HERE ---
    excel_data = convert_metrics_to_excel(individual_results)
    # Construct a dynamic filename
    thread_subject = get_clean_subject(individual_results[0].get('email', {}).get('subject', 'Thread'))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"evaluation_metrics_{thread_subject[:30].replace(' ', '_')}_{timestamp}.xlsx"
    
    st.download_button(
        label="📥 Download Complete Results Excel",
        data=excel_data,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("---") # Add a separator
    # --- END DOWNLOAD BUTTON ---
    
    # --- START DASHBOARD DISPLAY ---
    st.subheader("Performance Dashboard")
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
        
    # Display overall review 
    st.subheader("Overall Thread Review")
    with st.spinner("Generating overall thread review..."):
        overall_review_text = generate_overall_thread_review(individual_results)
        st.markdown(overall_review_text)

# Helper function to ensure results are properly saved for Results tab
def ensure_results_tab_works(results):
    """Ensure results are properly saved for the Results tab"""
    print("\n=== Debug: ensure_results_tab_works ===")
    print(f"Input results type: {type(results)}")
    print(f"Input results length: {len(results) if results else 0}")
    
    if not results:
        print("Warning: No results to save")
        return False
    
    try:
        # Deep copy the results to prevent reference issues
        import copy
        results_copy = copy.deepcopy(results)
        
        # Verify each result has required fields
        for i, result in enumerate(results_copy):
            print(f"\nVerifying result {i+1}:")
            
            # Check if the result has all required fields
            required_fields = ['email', 'ai_output', 'metrics']
            missing_fields = [field for field in required_fields if field not in result]
            
            if missing_fields:
                print(f"Warning: Result {i+1} is missing fields: {missing_fields}")
                # Add placeholder values for missing fields
                for field in missing_fields:
                    if field == 'email':
                        result['email'] = {'subject': 'Unknown Subject', 'id': f'email_{i}'}
                    elif field == 'ai_output':
                        result['ai_output'] = {'error': 'No AI output available'}
                    elif field == 'metrics':
                        result['metrics'] = []
            
            # Ensure the email index is set
            if 'email_index' not in result:
                result['email_index'] = i + 1
                print(f"Added missing email_index: {i+1}")
        
        # Store results in session state
        st.session_state.individual_results = results_copy
        st.session_state.has_results = True
        st.session_state.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Verify the save worked
        print("\nVerification after save:")
        print(f"Keys in session state: {list(st.session_state.keys())}")
        print(f"'individual_results' in session state: {'individual_results' in st.session_state}")
        print(f"Type of stored results: {type(st.session_state.individual_results)}")
        print(f"Length of stored results: {len(st.session_state.individual_results)}")
        
        if len(st.session_state.individual_results) != len(results):
            print("Warning: Stored results length doesn't match input length")
            return False
            
        # Print first result details for verification
        if st.session_state.individual_results:
            first_result = st.session_state.individual_results[0]
            print("\nFirst result verification:")
            print(f"Has email: {'email' in first_result}")
            print(f"Has metrics: {'metrics' in first_result}")
            print(f"Has ai_output: {'ai_output' in first_result}")
            
        return True
            
    except Exception as e:
        print(f"Error in ensure_results_tab_works: {str(e)}")
        import traceback
        traceback.print_exc()
        
        # Try one more time with a simpler approach
        try:
            print("Attempting fallback storage method")
            st.session_state["individual_results"] = results
            st.session_state.has_results = True
            print("Fallback storage completed")
            return True
        except Exception as e2:
            print(f"Fallback storage also failed: {str(e2)}")
            return False

# Helper function to format input data
def format_input_data(email):
    """
    Format email data into the required structure exactly matching the example format
    
    Args:
        email: Email dictionary from Graph API
        
    Returns:
        Dictionary with formatted email data matching the exact required format
    """
    try:
        # Extract basic info first
        mail_id = email.get("id", "")
        thread_id = email.get("conversationId", "")
        
        # Extract sender information
        sender = email.get("from", {}).get("emailAddress", {})
        sender_email = sender.get("address", "unknown@example.com")
        
        # Get received time
        received_time = email.get("receivedDateTime", "")
        
        # Get email body
        body = email.get("body", {})
        content_type = body.get("contentType", "").lower()
        content = body.get("content", "")
        
        # Clean the content to remove metadata
        cleaned_content = clean_email_content_remove_metadata(content)
        
        # Construct object exactly matching the format shown in example
        formatted_data = {
            "mail_id": mail_id,
            "file_name": [],
            "email": sender_email,
            "mail_time": received_time,
            "body_type": "plain",
            "mail_body": cleaned_content,
            "thread_id": thread_id,
            "mail_summary": ""
        }
        
        print(f"Formatted Input Data: {json.dumps(formatted_data, indent=2)}")
        return formatted_data
        
    except Exception as e:
        print(f"Error in format_input_data: {str(e)}")
        # Return minimal valid structure in correct format
        return {
            "mail_id": email.get("id", "unknown"),
            "file_name": [],
            "email": email.get("from", {}).get("emailAddress", {}).get("address", "unknown@example.com"),
            "mail_time": email.get("receivedDateTime", ""),
            "body_type": "plain",
            "mail_body": "",
            "thread_id": email.get("conversationId", "unknown"),
            "mail_summary": ""
        }

def map_ai_output(ai_result, email):
    """
    Map AI output to match the exact format shown in the example
    
    Args:
        ai_result: Result from AI service (Gemini or Dwellworks API)
        email: Original email dictionary
        
    Returns:
        Mapped output dictionary with standardized field names in exactly the required format
    """
    # Create output with default values in exactly the required format
    output = {
        "Sentiment analysis": "red",  # Only red/green are valid values
        "overall_sentiment_analysis": "neutral", 
        "feature": "no feature",
        "category": "no category",
        "Summary": "No summary available",
        "Events": [],
        "mail_id": email.get("id", ""),
        "thread_id": email.get("conversationId", "")
    }
    
    # If ai_result is not a dict, return defaults
    if not isinstance(ai_result, dict):
        print("AI result is not a dictionary, using default values")
        return output
    
    # Check for error in result
    if "error" in ai_result:
        print(f"Error in AI result: {ai_result['error']}")
        output["error"] = ai_result["error"]
        return output
    
    # Map sentiment analysis fields - handle all possible variations
    # For Sentiment analysis, only red/green are valid values
    sentiment_value = None
    if "Sentiment analysis" in ai_result:
        sentiment_value = ai_result["Sentiment analysis"]
    elif "sentiment_analysis" in ai_result:
        sentiment_value = ai_result["sentiment_analysis"]
    elif "Sentiment Analysis" in ai_result:
        sentiment_value = ai_result["Sentiment Analysis"]
    elif "sentiment" in ai_result:
        sentiment_value = ai_result["sentiment"]
    
    # Standardize sentiment value to only red/green
    if sentiment_value:
        if sentiment_value.lower() in ["positive", "green"]:
            output["Sentiment analysis"] = "green"
        else:
            # Default to red for any non-positive sentiment
            output["Sentiment analysis"] = "red"
    
    # Map overall sentiment - ensure it's positive/negative/neutral
    overall_sentiment = None
    if "overall_sentiment_analysis" in ai_result:
        overall_sentiment = ai_result["overall_sentiment_analysis"]
    elif "overall_sentiment" in ai_result:
        overall_sentiment = ai_result["overall_sentiment"]
    elif "sentiment_overall" in ai_result:
        overall_sentiment = ai_result["sentiment_overall"]
    
    # Standardize overall sentiment to positive/negative/neutral
    if overall_sentiment:
        if overall_sentiment.lower() in ["positive", "green"]:
            output["overall_sentiment_analysis"] = "positive"
        elif overall_sentiment.lower() in ["negative", "red"]:
            output["overall_sentiment_analysis"] = "negative"
        else:
            output["overall_sentiment_analysis"] = "neutral"
    
    # Map feature and category
    if "feature" in ai_result:
        output["feature"] = ai_result["feature"]
    elif "Feature" in ai_result:
        output["feature"] = ai_result["Feature"]
    
    if "category" in ai_result:
        output["category"] = ai_result["category"]
    elif "Category" in ai_result:
        output["category"] = ai_result["Category"]
    
    # Handle summary fields with different potential names
    if "Summary" in ai_result:
        output["Summary"] = ai_result["Summary"]
    elif "summary" in ai_result:
        output["Summary"] = ai_result["summary"]
    elif "email_summarization" in ai_result:
        output["Summary"] = ai_result["email_summarization"]
    elif "email_summary" in ai_result:
        output["Summary"] = ai_result["email_summary"]
    
    # Handle events with different potential formats
    events_list = []
    if "Events" in ai_result and ai_result["Events"]:
        events_list = ai_result["Events"] if isinstance(ai_result["Events"], list) else [ai_result["Events"]]
    elif "events" in ai_result and ai_result["events"]:
        events_list = ai_result["events"] if isinstance(ai_result["events"], list) else [ai_result["events"]]
    elif "events_summarization" in ai_result and ai_result["events_summarization"]:
        events_list = ai_result["events_summarization"] if isinstance(ai_result["events_summarization"], list) else [ai_result["events_summarization"]]
    
    # Standardize events to match the exact format required
    standardized_events = []
    for event in events_list:
        if isinstance(event, dict):
            standard_event = {
                "Event name": event.get("Event name", event.get("name", event.get("event_name", event.get("event", "Unknown event")))),
                "Date": event.get("Date", event.get("date", None)),
                "Time": event.get("Time", event.get("time", None)),
                "Property Type": event.get("Property Type", event.get("property_type", event.get("property", None))),
                "Agent Name": event.get("Agent Name", event.get("agent_name", event.get("agent", None))),
                "Location": event.get("Location", event.get("location", None))
            }
            standardized_events.append(standard_event)
        elif isinstance(event, str) and event.strip():
            # Handle case where event is just a string
            standard_event = {
                "Event name": event,
                "Date": None,
                "Time": None,
                "Property Type": None,
                "Agent Name": None,
                "Location": None
            }
            standardized_events.append(standard_event)
    
    if standardized_events:
        output["Events"] = standardized_events
    
    # Always include mail_id and thread_id
    output["mail_id"] = email.get("id", "")
    output["thread_id"] = email.get("conversationId", "")
    
    print(f"Mapped AI output: {output}")
    return output

def generate_groundtruth(email, previous_summary=None):
    """Generate groundtruth data matching the exact format provided in the example"""
    try:
        # Get current Gemini API key from session state
        api_key = st.session_state.get("gemini_api_key", "AIzaSyBMVP5wfR0R6LBLP_Tbbiaiudnaccau2IA")
        
        # Extract necessary information
        mail_id = email.get("id", "")
        thread_id = email.get("conversationId", "")
        body = email.get("body", {}).get("content", "")
        subject = email.get("subject", "")
        
        # Clean the content
        cleaned_content = clean_email_content_remove_metadata(body)
        
        # Create basic groundtruth structure first
        basic_groundtruth = {
            "Sentiment analysis": "red",  # Default to red as neutral is not valid for this field
            "overall_sentiment_analysis": "neutral",  # Can be positive, negative, or neutral
            "feature": "No feature",
            "category": "No category",
            "Summary": "",
            "Events": [],
            "mail_id": mail_id,
            "thread_id": thread_id
        }
        
        # Try to use Gemini for better groundtruth generation
        if api_key:
            try:
                import google.generativeai as genai
                
                # Configure Gemini
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                # Generate a thoughtful summary in indirect speech style with 2-3 sentences
                prompt = f"""
                Generate a concise, accurate summary of this email in indirect speech (reported speech) style.
                
                INSTRUCTIONS:
                1. Use indirect speech format (e.g., "The sender stated that they would..." NOT "I will...")
                2. Focus on the main message and key details only
                3. Keep the summary to 2-3 sentences at most
                4. Use third person perspective throughout
                5. Maintain formal, objective tone
                6. NEVER use direct quotes from the email
                7. Do NOT truncate or cut off the summary
                8. Ensure the summary is complete and coherent
                
                EXAMPLE OF CORRECT INDIRECT SPEECH FORMAT:
                Email: "Hi John, I'll meet you tomorrow at 2pm at the coffee shop. I'm bringing the documents you requested. Best, Sarah"
                
                ✅ GOOD SUMMARY (indirect speech): 
                "The sender informed John that they would meet him at 2pm the following day at the coffee shop. They mentioned they would bring the requested documents."
                
                ❌ BAD SUMMARY (direct speech): 
                "I'll meet you tomorrow at 2pm at the coffee shop. I'm bringing the documents you requested."
                
                EMAIL TO SUMMARIZE:
                {cleaned_content}
                
                SUBJECT: {subject}
                
                PREVIOUS SUMMARY: 
                {previous_summary if previous_summary else "None"}
                
                Generate ONLY the 2-3 sentence summary in proper indirect speech format.
                """
                
                # Get response from Gemini
                response = model.generate_content(prompt)
                
                # Extract summary from response
                summary = response.text.strip()
                
                # Update the basic groundtruth with Gemini-generated data
                basic_groundtruth["Summary"] = summary
                
                # Analyze sentiment
                sentiment_prompt = f"""
                Analyze the sentiment of this email. 
                You MUST return EXACTLY ONE of these values:
                1. "red" - if the sentiment is negative
                2. "green" - if the sentiment is positive
                
                IMPORTANT: For this task, there is NO neutral option. You must classify as either "red" or "green".
                Return ONLY the word "red" or "green" with no other text.
                
                EMAIL:
                {cleaned_content}
                """
                
                sentiment_response = model.generate_content(sentiment_prompt)
                sentiment = sentiment_response.text.strip().lower()
                
                # Normalize sentiment value for 'Sentiment analysis' - only red or green allowed
                if "red" in sentiment or "negative" in sentiment:
                    basic_groundtruth["Sentiment analysis"] = "red"
                else:
                    basic_groundtruth["Sentiment analysis"] = "green"
                
                # Analyze overall sentiment separately
                overall_sentiment_prompt = f"""
                Analyze the overall sentiment of this email. 
                Return EXACTLY ONE of these values:
                1. "negative" - if the overall tone is negative
                2. "positive" - if the overall tone is positive
                3. "neutral" - if the overall tone is neutral
                
                Return ONLY one word: "negative", "positive", or "neutral" with no other text.
                
                EMAIL:
                {cleaned_content}
                """
                
                overall_sentiment_response = model.generate_content(overall_sentiment_prompt)
                overall_sentiment = overall_sentiment_response.text.strip().lower()
                
                # Normalize overall sentiment value
                if "negative" in overall_sentiment:
                    basic_groundtruth["overall_sentiment_analysis"] = "negative"
                elif "positive" in overall_sentiment:
                    basic_groundtruth["overall_sentiment_analysis"] = "positive"
                else:
                    basic_groundtruth["overall_sentiment_analysis"] = "neutral"
                    
                # Add feature and category analysis
                feature_prompt = f"""
                Determine if this email falls into any of these features:
                - "EMAIL -- DSC First Contact with EE Completed"
                - "EMAIL -- EE First Contact with DSC"
                - "EMAIL -- Phone Consultation Scheduled"
                - "EMAIL -- Phone Consultation Completed"
                - "No feature"
                
                Return ONLY the exact feature name that matches.
                
                EMAIL:
                {cleaned_content}
                """
                
                feature_response = model.generate_content(feature_prompt)
                feature = feature_response.text.strip()
                
                if any(f in feature for f in ["DSC First Contact", "EE First Contact", "Phone Consultation"]):
                    basic_groundtruth["feature"] = feature
                    basic_groundtruth["category"] = "Initial Service Milestones"
                else:
                    basic_groundtruth["feature"] = "No feature"
                    basic_groundtruth["category"] = "No category"
                    
                # Extract events if any
                event_prompt = f"""
                Extract any events mentioned in this email with their details in this exact JSON format:
                [
                  {{
                    "Event name": "event description",
                    "Date": "date mentioned",
                    "Time": "time mentioned",
                    "Location": "location mentioned",
                    "Property Type": "property type if mentioned",
                    "Agent Name": "agent name if mentioned"
                  }}
                ]
                
                If no events are mentioned, return an empty array: []
                
                EMAIL:
                {cleaned_content}
                """
                
                event_response = model.generate_content(event_prompt)
                try:
                    import json
                    events = json.loads(event_response.text.strip())
                    basic_groundtruth["Events"] = events
                except:
                    # Unable to parse JSON, keep empty array
                    basic_groundtruth["Events"] = []
                    
            except Exception as e:
                print(f"Error using Gemini for groundtruth: {str(e)}")
                # In case of error, generate a very basic summary directly from the content
                if cleaned_content:
                    # Generate a simple summary based on the first few sentences
                    import re
                    sentences = re.split(r'[.!?]+', cleaned_content)
                    sentences = [s.strip() for s in sentences if s.strip()]
                    
                    if sentences:
                        # Get the first 1-2 sentences only
                        content_preview = ". ".join(sentences[:2])
                        # Convert to indirect speech format
                        sender = email.get("sender", {}).get("emailAddress", {}).get("name", "The sender")
                        recipient = "the recipient"
                        
                        summary = f"{sender} wrote to {recipient} regarding {subject}. "
                        summary += f"They mentioned {content_preview}."
                        
                        # Ensure it doesn't exceed 2-3 sentences
                        if len(summary.split('.')) > 3:
                            summary = '. '.join(summary.split('.')[:3]) + '.'
                            
                        basic_groundtruth["Summary"] = summary.strip()
        
        return basic_groundtruth
        
    except Exception as e:
        print(f"Error in generate_groundtruth: {str(e)}")
        return {
            "Sentiment analysis": "red",  # Default to red as neutral is not valid here
            "overall_sentiment_analysis": "neutral",  # Can be positive, negative, or neutral
            "feature": "No feature",
            "category": "No category",
            "Summary": "Unable to generate summary due to an error.",
            "Events": [],
            "mail_id": email.get("id", ""),
            "thread_id": email.get("conversationId", "")
        }

def clean_email_content_with_gemini(email_body: str) -> str:
    """Use Gemini to extract just the core email content without metadata"""
    try:
        import google.generativeai as genai
        import re
        
        # First try basic cleaning without Gemini
        def basic_clean(content):
            # Remove HTML tags
            content = re.sub(r'<[^>]+>', ' ', content)
            # Remove multiple spaces and newlines
            content = re.sub(r'\s+', ' ', content)
            # Split into lines
            lines = content.strip().split('\n')
            cleaned_lines = []
            for line in lines:
                line = line.strip()
                # Skip empty lines and common metadata patterns
                if not line or any(pattern in line.lower() for pattern in [
                    "from:", "to:", "sent:", "date:", "subject:",
                    "caution:", "disclaimer:", "confidential",
                    "original message", "forwarded message"
                ]):
                    continue
                cleaned_lines.append(line)
            return "\n".join(cleaned_lines).strip()
        
        # Try basic cleaning first
        cleaned_content = basic_clean(email_body)
        
        # Only use Gemini if basic cleaning doesn't give good results
        if len(cleaned_content) > 50:  # If we have reasonable content from basic cleaning
            return cleaned_content
            
        # Try Gemini as a fallback
        try:
            # --- MODIFIED SECTION for Gemini API Key ---
            api_key_to_use = st.session_state.get("gemini_api_key")
            if not api_key_to_use:
                print("Error: Gemini API Key not configured in session state for clean_email_content_with_gemini.")
                cleaned_content = basic_clean(email_body) 
                return cleaned_content

            genai.configure(api_key=api_key_to_use)
            model = genai.GenerativeModel('gemini-1.5-pro')
            # --- END MODIFIED SECTION ---
            
            prompt = f"""
            Extract ONLY the essential email content from the following text. Remove ALL metadata, system information, disclaimers, and forwarded message headers.
            
            Rules:
            1. Keep ONLY: greeting, main message body, and signature
            2. Remove ALL: email headers, timestamps, disclaimers, forwarded message markers, system-generated text
            3. Remove any "From:", "To:", "Subject:", "Date:" lines
            4. Remove any "CAUTION:" or warning messages
            5. Remove any legal disclaimers or confidentiality notices
            6. Format the output as a clean email with just greeting, content, and signature
            
            Input email:
            {email_body}
            
            Return ONLY the cleaned content in this format:
            [greeting]
            [message body]
            [signature]
            """
            
            response = model.generate_content(prompt)
            gemini_cleaned = response.text.strip()
            
            # Use Gemini result only if it's better than basic cleaning
            if len(gemini_cleaned) > len(cleaned_content):
                return gemini_cleaned
            
        except Exception as e:
            print(f"Gemini cleaning failed, using basic cleaning: {str(e)}")
            
        return cleaned_content
            
    except Exception as e:
        print(f"Error cleaning email content: {str(e)}")
        # Fallback to basic cleaning if everything else fails
        return basic_clean(email_body)

def process_individual_email(email, ai_service, previous_summary=None):
    """Process an individual email and get AI analysis results in exact required format"""
    try:
        # Start timer
        start_time = time.time()
        print("\n=== Processing Individual Email ===")
        print(f"Email ID: {email.get('id', 'unknown')}")
        print(f"Has previous summary: {previous_summary is not None}")
        
        # Format input data for analysis in exact required format
        input_data = format_input_data(email)
        
        # Add previous summary to input data if available
        if previous_summary:
            input_data["mail_summary"] = previous_summary
            print(f"Added previous summary to input data: {previous_summary[:100]}...")
            
        print(f"Input data formatted successfully")
        
        # Get Gemini API key from session state
        gemini_api_key = st.session_state.get("gemini_api_key", "AIzaSyBMVP5wfR0R6LBLP_Tbbiaiudnaccau2IA")
        
        # Initialize AI service if not provided
        if not ai_service:
            from ai_service import AIService
            ai_service = AIService()
            print("Created new AIService instance")
        
        ai_output_source = None # Flag to track the source of AI result
        api_result = None

        # Try using the Dwellworks API
        try:
            print("Attempting to use Dwellworks API for analysis...")
            # Pass email as a list since analyze_email_thread expects a list
            # Also pass the previous summary to the API
            api_result = ai_service.analyze_email_thread([email], previous_summary=previous_summary)
            print("Successfully used Dwellworks API")
            print(f"Dwellworks API result type: {type(api_result)}")
            if isinstance(api_result, dict):
                print(f"Dwellworks API result keys: {list(api_result.keys())}")
            ai_output_source = "dwellworks"
        except Exception as api_error:
            print(f"Dwellworks API error: {str(api_error)}, falling back to Gemini")
            # Fall back to Gemini
            current_gemini_api_key_fallback = st.session_state.get("gemini_api_key")
            if not current_gemini_api_key_fallback:
                # This scenario should ideally be handled by AIService or raise a more specific error
                # For now, if key is missing, api_result will remain None or be an error structure from use_gemini_for_analysis
                print("ERROR: Gemini API Key not found in session state for fallback. Analysis may fail or use AIService internal handling.")
            else:
                 genai.configure(api_key=current_gemini_api_key_fallback) # Ensure genai is configured for this specific call if needed by AIService

            gemini_input = {
                "sender_name": email.get("from", {}).get("emailAddress", {}).get("name", "Unknown"),
                "sender_email": email.get("from", {}).get("emailAddress", {}).get("address", "unknown@example.com"),
                "recipients": [r.get("emailAddress", {}).get("address", "") for r in email.get("toRecipients", [])],
                "content": email.get("body", {}).get("content", ""),
                "subject": email.get("subject", ""),
                "sent_time": email.get("receivedDateTime", ""),
                "previous_context": previous_summary
            }
            
            api_result = ai_service.use_gemini_for_analysis([gemini_input], feature_set="real_estate")
            print("Successfully used Gemini fallback")
            print(f"Gemini fallback result type: {type(api_result)}")
            if isinstance(api_result, dict):
                print(f"Gemini fallback result keys: {list(api_result.keys())}")
            ai_output_source = "gemini_fallback"

        # Generate groundtruth in exact required format
        print("Generating groundtruth data...")
        groundtruth = generate_groundtruth(email, previous_summary)
        print(f"Groundtruth generated with keys: {list(groundtruth.keys()) if isinstance(groundtruth, dict) else 'not a dict'}")
        
        # Conditionally determine ai_output
        ai_output = None
        if ai_output_source == "dwellworks":
            print("Using Dwellworks API response as is for ai_output.")
            ai_output = api_result 
            # Ensure mail_id and thread_id are present for context, even in raw output
            if isinstance(ai_output, dict):
                ai_output.setdefault("mail_id", email.get("id", ""))
                ai_output.setdefault("thread_id", email.get("conversationId", ""))
            # If Dwellworks can return a list for a single email, this needs more robust handling.
            # Assuming api_result from analyze_email_thread for a single email is a dict.
        elif ai_output_source == "gemini_fallback":
            print("Mapping Gemini fallback response to output format for ai_output...")
            ai_output = map_ai_output(api_result, email)
        else:
            # This case should ideally not be reached if the try/except for API calls is exhaustive
            print("ERROR: AI output source is unknown. Attempting default mapping.")
            # Fallback to default mapping if source is somehow not set, or handle as an error state
            ai_output = map_ai_output(api_result, email) if api_result else map_ai_output({}, email) # Ensure map_ai_output gets a dict

        print(f"Final ai_output type: {type(ai_output)}")
        if isinstance(ai_output, dict):
            print(f"Final ai_output keys: {list(ai_output.keys())}")

        # Create result with all necessary fields in correct format
        result = {
            "email": email,  # Keep original email for reference
            "input_data": input_data,  # Input in exact format
            "ai_output": ai_output,  # AI output in exact format
            "groundtruth": groundtruth,  # Groundtruth in exact format
            "processing_time": time.time() - start_time,
            "previous_summary": previous_summary  # Store the previous summary used
        }
        
        # Get email content for evaluation
        email_content = ""
        if 'email' in result and 'body' in result['email'] and 'content' in result['email']['body']:
            email_content = clean_email_content(result['email']['body']['content'])
        
        # Use LLM-based evaluation
        print("Evaluating AI output using LLM...")
        metrics = evaluate_with_llm(ai_output, groundtruth, email_content)
        print(f"LLM evaluation complete: generated {len(metrics)} metrics")
        
        # --- START: Ensure raw Dwellworks/mapped Gemini events are in the metric's 'AI Value' for UI display --- 
        # This is to make sure the UI displays the events as received from ai_service.py (for Dwellworks)
        # or as mapped (for Gemini fallback), if evaluate_with_llm doesn't populate it this way.
        if isinstance(ai_output, dict) and "Events" in ai_output:
            actual_events_in_ai_output = ai_output.get("Events") # This is the list of events
            if actual_events_in_ai_output is not None: # Check if there are events to inject
                event_metric_found_and_updated = False
                for metric_item in metrics:
                    metric_field_name = metric_item.get("Metric", metric_item.get("Field", "")).lower()
                    # Target metric names used in Tab 2 for event display
                    if metric_field_name in ['events', 'event match', 'events_count', 'event detection']:
                        # Preserve original explanations and status if possible, only override AI Value for display
                        if metric_item.get("AI Value") != actual_events_in_ai_output:
                            print(f"DEBUG: Forcing 'AI Value' of metric '{metric_field_name}' to events from ai_output.")
                            metric_item["AI Value"] = actual_events_in_ai_output
                        else:
                            print(f"DEBUG: Metric '{metric_field_name}' 'AI Value' already matches events from ai_output.")
                        event_metric_found_and_updated = True
                        break # Assuming one primary event metric to update for display
                
                if not event_metric_found_and_updated:
                    # If evaluate_with_llm produced no event metric at all, but ai_output had events.
                    # This situation is trickier as we'd be missing status, explanations etc.
                    # For now, we only modify an existing event metric.
                    print(f"DEBUG: ai_output had 'Events', but no corresponding metric found in LLM output to update its 'AI Value'. Events may not display if not in a metric.")
        # --- END: Ensure raw events in metric --- 

        # Check if metrics is empty or None and log a warning
        if not metrics:
            print("WARNING: LLM evaluation returned empty metrics")
            # Generate fallback metrics
            metrics = generate_fallback_metrics(ai_output, groundtruth)
            print(f"Generated {len(metrics)} fallback metrics")
        
        result["metrics"] = metrics
        print("Metrics successfully added to result")
        
        # Debug: Print what's in the result object
        print("\nResult object contains:")
        for key in result:
            if key == "email":
                print(f"  email: [email object]")
            elif key == "metrics":
                print(f"  metrics: {len(result['metrics'])} items")
            else:
                print(f"  {key}: {type(result[key])}")
        
        return result
        
    except Exception as e:
        print(f"Error in process_individual_email: {str(e)}")
        import traceback
        traceback.print_exc()
        
        # Even on error, provide a valid result with minimal fields
        return {
            "error": str(e),
            "email": email,
            "input_data": format_input_data(email) if email else None,
            "ai_output": {
                "Sentiment analysis": "red",  # Changed from "yellow" to match valid values
                "overall_sentiment_analysis": "neutral",
                "feature": "no feature",
                "category": "no category",
                "Summary": f"Error during analysis: {str(e)}",
                "Events": [],
                "mail_id": email.get("id", "") if email else "",
                "thread_id": email.get("conversationId", "") if email else ""
            },
            "groundtruth": generate_groundtruth(email, previous_summary) if 'generate_groundtruth' in globals() and email else None,
            "metrics": [{
                "Field": "Error",
                "AI Value": f"Processing error: {str(e)}",
                "Ground Truth": "N/A",
                "Status": "Fail",
                "Evidence": "Error occurred during processing",
                "Explanation": "This is a placeholder due to an error."
            }]
        }

def force_extract_email_content(content):
    """Force extraction of email content when basic cleaning fails."""
    import re
    from bs4 import BeautifulSoup
    
    try:
        # First try to parse as HTML
        soup = BeautifulSoup(content, 'html.parser')
        
        # Remove script and style elements
        for script in soup(["script", "style"]):
            script.decompose()
        
        # Try to find main content
        main_content = None
        
        # Look for common email content containers
        for div in soup.find_all('div'):
            if div.get('class'):
                classes = ' '.join(div.get('class')).lower()
                if any(term in classes for term in ['content', 'body', 'message', 'email']):
                    main_content = div
                    break
        
        if not main_content:
            # If no specific content div found, use the longest text block
            text_blocks = []
            for tag in soup.find_all(['p', 'div']):
                text = tag.get_text(strip=True)
                if len(text) > 50:  # Only consider blocks with substantial content
                    text_blocks.append(text)
            
            if text_blocks:
                main_content = max(text_blocks, key=len)
            else:
                main_content = soup.get_text(separator=' ', strip=True)
        
        # Get text content
        if isinstance(main_content, str):
            text = main_content
        else:
            text = main_content.get_text(separator=' ', strip=True)
        
        # Clean up the text
        text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces
        text = re.sub(r'[^\x00-\x7F]+', '', text)  # Remove non-ASCII
        text = re.sub(r'From:.*?Subject:', '', text)  # Remove email headers
        text = re.sub(r'On.*?wrote:', '', text)  # Remove reply headers
        text = re.sub(r'________________________________', '', text)  # Remove separators
        
        return text.strip()
    except Exception as e:
        print(f"Error in force extraction: {str(e)}")
        return content

def clean_email_content(content):
    """Clean email content by removing metadata and formatting while preserving property details"""
    if not content:
        return ""
    
    # Use the improved metadata removal approach
    clean_content = clean_email_content_remove_metadata(content)
    
    # If we still have a lot of content, use the more thorough cleaning approach
    if len(clean_content) > 500:
        from bs4 import BeautifulSoup
        import re
        
        try:
            # Check if content is HTML
            if "<html" in clean_content.lower() or ("<" in clean_content and ">" in clean_content):
                # Parse with BeautifulSoup
                soup = BeautifulSoup(clean_content, 'html.parser')
                
                # Remove script and style elements
                for element in soup(["script", "style"]):
                    element.decompose()
                
                # Get text content
                clean_content = soup.get_text(separator=' ', strip=True)
            
            # Remove excessive whitespace and normalize spaces
            clean_content = re.sub(r'\s+', ' ', clean_content).strip()
            
            # Preserve property details by ensuring addresses and dates are kept intact
            property_patterns = [
                r'\b\d+\s+[A-Za-z\s]+(?:Road|Rd|Street|St|Avenue|Ave|Boulevard|Blvd|Drive|Dr|Lane|Ln|Way|Place|Pl|Court|Ct|Terrace|Ter|Trail|Trl|Park|Circle|Cir)\b',
                r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?,\s+\d{4}\b',
                r'\b\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)\b'
            ]
            
            # Ensure property details are preserved
            original_content = content
            for pattern in property_patterns:
                for match in re.finditer(pattern, original_content, re.IGNORECASE):
                    match_text = match.group(0)
                    if match_text not in clean_content:
                        clean_content += f" {match_text}"
        
        except Exception as e:
            print(f"Error in advanced cleaning: {str(e)}")
            # If advanced cleaning fails, return content with just metadata removed
            return clean_content
    
    return clean_content

def basic_clean_email_content(content):
    """Basic cleaning of email content."""
    import re
    from bs4 import BeautifulSoup
    
    try:
        # Remove HTML
        soup = BeautifulSoup(content, 'html.parser')
        text = soup.get_text(separator=' ', strip=True)
        
        # Basic cleaning
        text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces with single space
        text = re.sub(r'[^\x00-\x7F]+', '', text)  # Remove non-ASCII characters
        
        return text.strip()
    except Exception as e:
        print(f"Error in basic cleaning: {str(e)}")
        return content

def clean_email_content_remove_metadata(content):
    """
    Clean email content by removing everything after metadata markers.
    This helps remove signatures, disclaimers, forwarded content, etc.
    
    Args:
        content: The original email content
        
    Returns:
        Cleaned email content with metadata removed
    """
    import re
    from bs4 import BeautifulSoup
    
    if not content:
        return ""
    
    # First check if this is HTML content
    if "<html" in content.lower() or ("<" in content and ">" in content):
        # Parse HTML
        try:
            soup = BeautifulSoup(content, 'html.parser')
            content = soup.get_text(separator='\n', strip=True)
        except Exception as e:
            print(f"Error parsing HTML: {str(e)}")
    
    # Split content into lines for better processing
    lines = content.split('\n')
    cleaned_lines = []
    found_metadata = False
    
    for line in lines:
        line = line.strip()
        
        # Skip empty lines
        if not line:
            continue
            
        # Check for exact "From:" match (case-insensitive)
        if re.match(r'^[Ff][Rr][Oo][Mm]:', line):
            found_metadata = True
            break
            
        # Other metadata markers to check
        metadata_markers = [
            r'^[Ss][Ee][Nn][Tt]:',
            r'^[Tt][Oo]:',
            r'^[Ss][Uu][Bb][Jj][Ee][Cc][Tt]:',
            r'^[Dd][Aa][Tt][Ee]:',
            r'^[Cc][Cc]:',
            r'^[Bb][Cc][Cc]:',
            r'^>{2,}',  # Multiple > characters indicating quoted text
            r'^-{3,}',  # Three or more hyphens
            r'^_{3,}',  # Three or more underscores
            r'^[*]{3,}',  # Three or more asterisks
            r'^[Cc][Aa][Uu][Tt][Ii][Oo][Nn]:',
            r'^[Dd][Ii][Ss][Cc][Ll][Aa][Ii][Mm][Ee][Rr]:',
            r'^[Cc][Oo][Nn][Ff][Ii][Dd][Ee][Nn][Tt][Ii][Aa][Ll]',
            r'^[Oo][Rr][Ii][Gg][Ii][Nn][Aa][Ll] [Mm][Ee][Ss][Ss][Aa][Gg][Ee]',
            r'^[Oo][Nn] .+wrote:$',  # "On ... wrote:" pattern
            r'^[Bb]est [Rr]egards',
            r'^[Rr]egards,',
            r'^[Ss]incerely,',
            r'^[Tt]hank(s| you),?$',
            r'^[Cc]heers,?$'
        ]
        
        if any(re.match(pattern, line) for pattern in metadata_markers):
            found_metadata = True
            break
            
        cleaned_lines.append(line)
    
    # Join the cleaned lines
    cleaned_content = '\n'.join(cleaned_lines).strip()
    
    # If no metadata markers found, return the original content
    if not found_metadata and not cleaned_content:
        return content.strip()
        
    return cleaned_content

def display_email_content(content):
    """Display email content with proper formatting and metadata removal"""
    if not content or len(content.strip()) == 0:
        st.markdown(
            '<div class="empty-content">No email content available</div>',
            unsafe_allow_html=True
        )
    else:
        # Extract content before processing
        clean_content = content
        
        # Apply more thorough metadata removal
        clean_content = clean_email_content_remove_metadata(clean_content)
        
        # Clean up any remaining HTML
        import re
        from bs4 import BeautifulSoup
        
        # Handle HTML content if still present
        if "<html" in clean_content.lower() or ("<" in clean_content and ">" in clean_content):
            try:
                # Extract body content
                body_match = re.search(r'<body[^>]*>(.*?)</body>', clean_content, re.DOTALL | re.IGNORECASE)
                if body_match:
                    clean_content = body_match.group(1)
                # Remove HTML tags
                soup = BeautifulSoup(clean_content, 'html.parser')
                clean_content = soup.get_text(separator=' ', strip=True)
            except Exception as e:
                print(f"Error cleaning HTML content: {str(e)}")
        
        # Last pass to clean whitespace
        clean_content = re.sub(r'\s+', ' ', clean_content).strip()
        
        # Display the cleaned content
        st.markdown(
            f'<div class="email-content">{clean_content}</div>',
            unsafe_allow_html=True
        )

# Import local modules
from graph_api_client import GraphAPIClient
from vector_store import VectorStore
from ai_service import AIService

# Initialize the vector store right at the top
os.makedirs("email_data", exist_ok=True)
vector_store = VectorStore("email_data")

# Set page config
st.set_page_config(
    page_title="Email AI Automation",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items=None
)

# CSS for styling
st.markdown("""
<style>
    /* Hide Streamlit toolbar elements */
    .stToolbar {
        display: none !important;
    }
    .viewerBadge_container__1QSob {
        display: none !important;
    }
    .viewerBadge_link__1S137 {
        display: none !important;
    }
    .viewerBadge_text__1JaDK {
        display: none !important;
    }
    header[data-testid="stHeader"] {
        display: none !important;
    }
    
    /* Remove Streamlit branding and menu */
    #MainMenu {
        display: none !important;
    }
    footer {
        display: none !important;
    }
    
    .main .block-container {
        padding-top: 1rem;
    }
    
    /* App title */
    .app-title {
        font-size: 2.3rem;
        font-weight: 600;
        color: white;
        margin-bottom: 0.5rem;
        padding-left: 1rem;
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background-color: #1E1E1E;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        padding: 0.5rem;
        background-color: #1A1A1A;
        border-radius: 4px;
        margin-bottom: 0.2rem; /* MODIFIED: Reduced bottom margin to shrink space below tabs */
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        background-color: #2D2D2D !important;
        border-radius: 4px !important;
        padding: 8px 16px;
        color: #CCC;
        font-weight: 400;
        border: none !important;
        margin-right: 4px;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #4A4A4A !important;
        color: white !important;
        font-weight: 600;
    }
    
    /* Header styling */
    h1, h2, h3 {
        color: white;
        font-weight: 600;
    }
    
    /* Filter section styling */
    .filter-container {
        background-color: #1E1E1E;
        padding: 0.5rem 1.5rem 1.5rem 1.5rem; /* MODIFIED: Reduced top padding */
        border-radius: 4px;
        margin-bottom: 1.5rem;
        border: 1px solid #333;
    }
    
    /* Button styling */
    .stButton button {
        background-color: #3A3A3A;
        color: white;
        border: none;
        padding: 0.5rem 1rem;
        border-radius: 4px;
    }
    
    .stButton button:hover {
        background-color: #4A4A4A;
    }
    
    /* Primary button */
    .stButton button[kind="primary"] {
        background-color: #FF4B4B;
    }
    
    .stButton button[kind="primary"]:hover {
        background-color: #FF6B6B;
    }
    
    /* Error message styling */
    .stAlert {
        border-radius: 4px;
    }
    
    /* Results table */
    .results-table th {
        background-color: #262730;
        color: white;
        text-align: left;
        padding: 12px;
        border: 1px solid #444;
    }
    
    .results-table td {
        padding: 12px;
        border: 1px solid #444;
        background-color: #1A1A1A;
    }
    
    /* Input fields */
    div[data-baseweb="base-input"] {
        background-color: #111;
        border-radius: 4px;
    }
    
    div[data-baseweb="base-input"] input {
        color: #DDD;
    }
    
    /* Toggle buttons */
    div[data-testid="stExpander"] div[role="button"] p {
        font-size: 1.1rem;
        font-weight: 600;
    }
    
    /* Remove padding around info messages */
    div.stAlert {
        padding: 0.5rem;
    }
    
    /* Column styling */
    [data-testid="column"] {
        background-color: #1E1E1E;
        padding: 1rem;
        border-radius: 4px;
        border: 1px solid #333;
        margin: 0.25rem;
    }
    
    /* Email content styling */
    .email-content {
        background-color: #2D2D2D !important;
        border: 1px solid #444 !important;
        border-radius: 4px !important;
        padding: 20px !important;
        margin: 10px 0 !important;
        max-height: 300px !important;
        overflow-y: auto !important;
        white-space: pre-wrap !important;
        word-wrap: break-word !important;
        font-family: monospace !important;
        color: #E0E0E0 !important;
        font-size: 14px !important;
        line-height: 1.5 !important;
    }
    
    /* Empty email content placeholder */
    .empty-content {
        background-color: #222 !important;
        border: 1px dashed #444 !important;
        border-radius: 4px !important;
        padding: 20px !important;
        margin: 10px 0 !important;
        text-align: center !important;
        color: #888 !important;
        font-style: italic !important;
    }
    
    /* JSON content styling */
    div.element-container div.stJson {
        background-color: #2D2D2D !important;
        border: 1px solid #444 !important;
        border-radius: 4px !important;
        padding: 15px !important;
        margin: 5px 0 !important;
        max-width: 100% !important;
        overflow-x: hidden !important;
        word-wrap: break-word !important;
        word-break: break-all !important;
        white-space: pre-wrap !important;
        font-family: monospace !important;
        font-size: 12px !important;
        line-height: 1.4 !important;
        color: #E0E0E0 !important;
    }
    
    /* Ensure long IDs don't overflow */
    .stJson {
        max-width: 100% !important;
        overflow-wrap: break-word !important;
        word-wrap: break-word !important;
        word-break: break-all !important;
    }
    
    /* Fix for long strings in JSON data */
    .react-json-view .string-value {
        word-break: break-all !important;
        white-space: normal !important;
        max-width: 100% !important;
    }
    
    /* Style JSON keys */
    .json-key {
        color: #88CCF1 !important;
    }
    
    /* Style JSON values */
    .json-value {
        color: #B5CEA8 !important;
    }
    
    /* Allow line breaks in table cells */
    .dataframe td {
        white-space: normal !important;
        word-break: break-word !important;
    }
    
    /* Improve display of IDs in tables */
    .dataframe td:has(br) {
        line-height: 1.5 !important;
    }
    
    /* Fix Streamlit JSON formatting to contain long IDs */
    pre {
        white-space: pre-wrap !important;       /* css-3 */
        white-space: -moz-pre-wrap !important;  /* Mozilla */
        white-space: -pre-wrap !important;      /* Opera 4-6 */
        white-space: -o-pre-wrap !important;    /* Opera 7 */
        word-wrap: break-word !important;       /* Internet Explorer 5.5+ */
        word-break: break-all !important;
        overflow-wrap: break-word !important;
    }
    
    /* Add specific style for mail_id and thread_id in JSON */
    .mail-id-display, .thread-id-display {
        max-width: 100% !important;
        word-wrap: break-word !important;
        word-break: break-all !important;
        font-size: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# Main title
st.markdown('<div class="app-title">Email AI Automation</div>', unsafe_allow_html=True)

# ------------------------------------------------------------
# Sidebar - Settings and help section
# ------------------------------------------------------------

with st.sidebar:
    st.title("Email Analysis Pipeline")
    st.markdown("---")
    
    # Gemini API Key Section
    with st.expander("🔑 Gemini API Key", expanded=True):
        st.session_state.gemini_api_key = st.text_input(
            "Gemini API Key", 
            value=st.session_state.get('gemini_api_key', ""), 
            type="password"
        )
        st.markdown("""
        <div style="background-color: #2D2D2D; padding: 10px; border-radius: 4px;">
            <p style="color: #E0E0E0; font-size: 14px;">
                <strong>Usage:</strong>
                <ol>
                    <li>Get your Gemini API key from <a href="https://makersuite.google.com/app/apikey" target="_blank" style="color: #88CCF1;">Google MakerSuite</a></li>
                    <li>Enter the key here and click "Save API Key"</li>
                </ol>
            </p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Save API Key"):
            st.success("Gemini API Key saved successfully!")
    
    # Outlook Credentials Section - Only show when Outlook option is selected
    if st.session_state.email_source == "Fetch emails from Outlook":
        with st.expander("📧 Outlook Credentials", expanded=True):
            st.session_state.ms_graph_client_id = st.text_input(
                "MS Graph Client ID", 
                value=st.session_state.get('ms_graph_client_id', ""),
                type="password"
            )
            st.session_state.ms_graph_client_secret = st.text_input(
                "MS Graph Client Secret", 
                value=st.session_state.get('ms_graph_client_secret', ""), 
                type="password"
            )
            st.session_state.ms_graph_tenant_id = st.text_input(
                "MS Graph Tenant ID", 
                value=st.session_state.get('ms_graph_tenant_id', ""), 
                type="password"
            )
            st.session_state.user_email_address = st.text_input(
                "User Email", 
                value=st.session_state.get('user_email_address', "")
            )
            if st.button("Save Outlook Settings"):
                st.success("Outlook settings saved successfully!")
    
    # Help section
    with st.expander("📚 How to Use This Tool", expanded=False):
        st.markdown("""
        ### Email Analysis Tool
        
        **Getting Started:**
        1. Enter your Gemini API key in the sidebar section above
        2. Choose your email source using the radio buttons:
           - **Upload Excel:** Upload a file with an 'Email Content' column
           - **Outlook:** Configure Outlook credentials in the sidebar first
        
        **Processing Workflow:**
        - For Excel: Upload a file → Click "Process Uploaded Excel"
        - For Outlook: Fetch threads → Select a thread → Click "Process Thread"
        
        **Results:**
        - View performance metrics in the dashboard
        - Check the overall thread review
        - Download a complete Excel report with all analysis results
        
        **Excel Report Contains:**
        - Evaluation Metrics sheet with detailed analysis
        - Overall Review sheet with 5-point summary
        - Email Data sheet with raw input/output/groundtruth data
        """)
    
    st.markdown("---")
    st.caption("Email AI Automation v1.0")

# ------------------------------------------------------------
# Single Tab Layout - All functionality in one tab
# ------------------------------------------------------------

# --- START: Email Source Selection ---
# Set up default selection if needed
if "email_source" not in st.session_state:
    st.session_state.email_source = "Upload emails as Excel"  # Default option matches session state init
    
# Create a callback function to handle radio button changes
def handle_radio_change():
    # This function will be called when the radio selection changes
    pass  # We don't need to do anything here, the rerun below will handle the UI update

# Display radio button for selection
selected_option = st.radio(
    "Select Email Source:",
    options=["Upload emails as Excel", "Fetch emails from Outlook"],
    index=0 if st.session_state.email_source == "Upload emails as Excel" else 1,
    key="email_source_radio_unique",  # Add unique key to prevent state conflicts
    on_change=handle_radio_change,  # Add callback function
    horizontal=True
)

# Update session state and force rerun if selection changed
if selected_option != st.session_state.email_source:
    st.session_state.email_source = selected_option
    st.experimental_rerun()  # Force a rerun to update the UI

# Separator
st.markdown("---")

# --- Excel Upload Section ---
if selected_option == "Upload emails as Excel":
    st.subheader("Process Emails from Excel File")
    
    # Instructions for users
    st.info("Please upload an Excel file containing email content. The file must have a column named 'Email Content'.")
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Upload an Excel file with an 'Email Content' column (.xlsx, .xls)",
        type=["xlsx", "xls"],
        key="excel_uploader"
    )
    
    if uploaded_file is not None:
        # Show preview of file before processing
        try:
            preview_df = pd.read_excel(uploaded_file, nrows=3)
            # st.write("Preview of uploaded file:")
            # st.dataframe(preview_df)
            
            # Check if 'Email Content' column exists
            if "Email Content" not in preview_df.columns:
                st.error("❌ The Excel file does not contain a column named 'Email Content'. Please check your file format.")
                column_names = list(preview_df.columns)
                st.write(f"Available columns: {', '.join(column_names)}")
            else:
                st.success("✅ File uploaded successfully and 'Email Content' column found.")
                
                # Process button
                if st.button("Process Uploaded Excel", type="primary", key="process_excel_btn"):
                    st.session_state.individual_results = []
                    st.session_state.has_results = False
                    st.session_state.thread_structure = None
                    
                    with st.spinner("Processing uploaded Excel file..."):
                        try:
                            # Reset file pointer before reading again
                            uploaded_file.seek(0)
                            df = pd.read_excel(uploaded_file)
                            
                            if "Email Content" not in df.columns:
                                st.error("Excel file must contain a column named 'Email Content'. Please check the file and try again.")
                                st.stop()
                            
                            email_bodies_from_excel = df["Email Content"].dropna().astype(str).tolist()
                            
                            if not email_bodies_from_excel:
                                st.warning("No email content found in the 'Email Content' column or column is empty.")
                                st.stop()
                            
                            results_list = []
                            previous_summary = ""
                            
                            # Initialize AI service for processing
                            try:
                                ai_service = AIService()
                            except Exception as ai_init_error:
                                st.error(f"Failed to initialize AI service: {str(ai_init_error)}")
                                st.stop()
                            
                            progress_bar_excel = st.progress(0)
                            status_message = st.empty()
                            
                            for idx, email_body in enumerate(email_bodies_from_excel):
                                progress_excel = (idx + 1) / len(email_bodies_from_excel)
                                progress_bar_excel.progress(progress_excel, text=f"Processing email {idx + 1} of {len(email_bodies_from_excel)} from Excel...")
                                status_message.info(f"Processing email {idx + 1}...")
                                
                                current_time_iso = datetime.now().isoformat() + "Z"
                                mock_email_object = {
                                    "id": f"excel_email_{idx + 1}",
                                    "conversationId": "excel_upload_thread_01",
                                    "subject": f"Email {idx + 1} from Excel Upload",
                                    "body": {"content": email_body, "contentType": "text"},
                                    "from": {"emailAddress": {"name": "Excel Upload", "address": "excel@example.com"}},
                                    "sender": {"emailAddress": {"name": "Excel Upload", "address": "excel@example.com"}},
                                    "toRecipients": [],
                                    "ccRecipients": [],
                                    "bccRecipients": [],
                                    "receivedDateTime": current_time_iso,
                                    "inferenceClassification": "focused",
                                    "parentFolderId": "mock_folder_id",
                                    "isDraft": False,
                                    "isRead": True
                                }
                                
                                current_gemini_api_key_excel = st.session_state.get("gemini_api_key")
                                if current_gemini_api_key_excel:
                                    import google.generativeai as genai
                                    genai.configure(api_key=current_gemini_api_key_excel)
                                else:
                                    status_message.warning("No Gemini API key found. Processing may fail if API access is needed.")
                                
                                try:
                                    result = process_individual_email(mock_email_object, ai_service, previous_summary)
                                    
                                    if result:
                                        result["email_index"] = idx + 1
                                        results_list.append(result)
                                        if isinstance(result.get("ai_output"), dict) and "Summary" in result["ai_output"]:
                                            previous_summary = result["ai_output"]["Summary"]
                                        else:
                                            raw_ai_summary = result.get("ai_output", {}).get("Summary", result.get("ai_output", {}).get("summary"))
                                            if raw_ai_summary:
                                                previous_summary = raw_ai_summary
                                            else:
                                                previous_summary = ""
                                except Exception as email_process_error:
                                    status_message.error(f"Error processing email {idx + 1}: {str(email_process_error)}")
                                    # Continue with next email instead of stopping entirely
                            
                            progress_bar_excel.empty()
                            status_message.empty()
                            
                            if results_list:
                                ensure_results_tab_works(results_list)
                                if 'individual_results' in st.session_state and len(st.session_state.individual_results) > 0:
                                    st.success(f"Successfully analyzed {len(st.session_state.individual_results)} emails from the Excel file!")
                                    
                                    # --- Display Dashboard and Download Button ---
                                    display_results_dashboard()
                                else:
                                    st.error("Error: Results from Excel were processed but not stored properly. Please try again or check logs.")
                            else:
                                st.warning("No results were generated from the Excel file. This could be due to empty content or processing errors for all rows.")
                        
                        except Exception as e:
                            st.error(f"An error occurred while processing the Excel file: {str(e)}")
                            import traceback
                            print("Error details from Excel processing:")
                            print(traceback.format_exc())
        except Exception as e:
            st.error(f"Error previewing Excel file: {str(e)}")
    else:
        st.info("Please upload an Excel file to start processing.")
    
    # Sample excel format information
    with st.expander("ℹ️ Excel File Format Information"):
        st.markdown("""
        **Required Excel Format:**
        
        Your Excel file should contain at least one column named exactly "Email Content" with each cell containing the content of a different email.
        
        **Example:**
        
        | Email Content |
        |---------------|
        | Dear John, I am writing to confirm our meeting tomorrow at 2pm... |
        | Hi Team, Please find attached the latest project report... |
        
        Additional columns will be ignored during processing.
        """)
        
        # Add sample Excel download option
        sample_data = {'Email Content': [
            'Dear John, I am writing to confirm our meeting tomorrow at 2pm. Looking forward to discussing the project progress.',
            'Hi Team, Please find attached the latest project report. We have made significant progress on the Phase 1 deliverables.'
        ]}
        sample_df = pd.DataFrame(sample_data)
        sample_buffer = BytesIO()
        with pd.ExcelWriter(sample_buffer, engine='xlsxwriter') as writer:
            sample_df.to_excel(writer, index=False, sheet_name='Sample')
        
        st.download_button(
            label="Download Sample Excel File",
            data=sample_buffer.getvalue(),
            file_name="sample_email_content.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- Outlook Fetching Section ---
elif selected_option == "Fetch emails from Outlook":
    st.subheader("Outlook Email Processing")
    
    # Filters section
    st.markdown('<div class="filter-container">', unsafe_allow_html=True)
    st.subheader("Email Filters")
    
    col1, col2 = st.columns(2)
    
    with col1:
        from_filter = st.text_input("From (Sender Email)")
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
                client_id_to_use = st.session_state.get('ms_graph_client_id')
                client_secret_to_use = st.session_state.get('ms_graph_client_secret')
                tenant_id_to_use = st.session_state.get('ms_graph_tenant_id')
                
                if not all([client_id_to_use, client_secret_to_use, tenant_id_to_use]):
                    st.error("MS Graph API credentials are not fully configured in settings. Please check the sidebar.")
                    st.stop()
                graph_client = GraphAPIClient(client_id_to_use, client_secret_to_use, tenant_id_to_use)
                access_token = graph_client.get_access_token()
                
                if not access_token:
                    st.error('Failed to get access token. Check your credentials.')
                else:
                    user_email_to_use = st.session_state.get('user_email_address')
                    if not user_email_to_use:
                        st.error("User email is not configured in settings. Please check the sidebar.")
                        st.stop()
                    thread_list, error = fetch_threads(
                        graph_client,
                        user_email_to_use,
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
        for thread in st.session_state.threads:
            # Format as "[Subject] (Count: X emails)"
            label = f"{thread['subject']} (Count: {thread['message_count']} emails)"
            thread_labels[label] = thread
            thread_options.append(label)
        
        selected_thread_label = st.selectbox(
            "Select Email Thread to Process:",
            options=thread_options,
            index=0 if thread_options else None
        )
        
        if not selected_thread_label:
            st.info("No threads available. Please fetch emails first.")
        else:
            # Get the selected thread object
            selected_thread = thread_labels[selected_thread_label]
            
            # Show thread details in a nice formatted box
            st.markdown("### Thread Information")
            
            # Create two columns for thread info display
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(f"**Subject:** {selected_thread['subject']}")
                st.markdown(f"**Number of Emails:** {selected_thread['message_count']}")
            
            with col2:
                st.markdown(f"**Thread ID:** {selected_thread['thread_id']}")
                st.markdown(f"**Latest Activity:** {selected_thread['latest_message_date']}")
            
            # Process Thread button
            if st.button("Process Thread", type="primary"):
                with st.spinner("Loading and analyzing email thread..."):
                    try:
                        client_id_to_use_process = st.session_state.get('ms_graph_client_id')
                        client_secret_to_use_process = st.session_state.get('ms_graph_client_secret')
                        tenant_id_to_use_process = st.session_state.get('ms_graph_tenant_id')

                        if not all([client_id_to_use_process, client_secret_to_use_process, tenant_id_to_use_process]):
                            st.error("MS Graph API credentials are not fully configured in settings for thread processing. Please check the sidebar.")
                            st.stop()
                        graph_client = GraphAPIClient(client_id_to_use_process, client_secret_to_use_process, tenant_id_to_use_process)
                        access_token = graph_client.get_access_token()
                    
                        if not access_token:
                            st.error('Failed to get access token. Check your credentials.')
                        else:
                            current_user_email = st.session_state.get('user_email_address')
                            if not current_user_email:
                                st.error("Processing halted: User Email is not configured in settings. Please check the sidebar.")
                                st.stop() 

                            # Clear any existing results
                            st.session_state.individual_results = []
                            st.session_state.has_results = False
                            
                            # Get thread messages - this should now include all emails in the normalized thread
                            thread_emails = []
                            try:
                                # Get the thread ID of the selected thread
                                thread_id = selected_thread['thread_id']
                                print(f"Retrieving all emails for thread ID: {thread_id}")
                                
                                # Check if we already have message IDs
                                message_ids = selected_thread.get('messages', [])
                                if message_ids and isinstance(message_ids, list):
                                    print(f"Thread has {len(message_ids)} message IDs, fetching full messages")
                                    thread_emails = []
                                    
                                    # Fetch each message by ID
                                    for msg_id in message_ids:
                                        try:
                                            if isinstance(msg_id, str):
                                                user_email_for_fetch = st.session_state.get('user_email_address')
                                                if not user_email_for_fetch:
                                                    print("Error: User email not found in session state for _get_email_with_body")
                                                    continue 
                                                full_msg = graph_client._get_email_with_body(user_email_for_fetch, msg_id)
                                                if full_msg:
                                                    thread_emails.append(full_msg)
                                        except Exception as e:
                                            print(f"Error fetching message {msg_id}: {str(e)}")
                                    
                                    # Create a minimal thread structure
                                    if thread_emails:
                                        thread_structure = []
                                        # Sort by date
                                        thread_emails.sort(key=lambda e: e.get('receivedDateTime', ''))
                                        # Add first email as root
                                        if thread_emails:
                                            root = thread_emails[0].copy()
                                            root['replies'] = []
                                            thread_structure.append(root)
                                            # Add other emails as replies to root
                                            for email in thread_emails[1:]:
                                                root['replies'].append(email)
                                        
                                        # Store thread structure
                                        st.session_state.thread_structure = thread_structure
                                else:
                                    # Try using the improved fetch_thread_messages method 
                                    # if we don't have message IDs
                                    user_email_for_thread_msgs = st.session_state.get('user_email_address')
                                    if not user_email_for_thread_msgs:
                                        st.error("User email not configured for fetching thread messages.")
                                        st.stop()
                                    thread_result = graph_client.fetch_thread_messages(user_email_for_thread_msgs, thread_id)
                                    
                                    # Extract the messages and thread structure from the result
                                    if isinstance(thread_result, dict):
                                        thread_emails = thread_result.get('messages', [])
                                        thread_structure = thread_result.get('thread_structure', [])
                                        
                                        if thread_emails:
                                            print(f"Retrieved {len(thread_emails)} messages in thread")
                                            
                                            # Store thread structure for visualization
                                            st.session_state.thread_structure = thread_structure
                                        else:
                                            print("No emails found in the thread result")
                                    else:
                                        print(f"Unexpected result type from fetch_thread_messages: {type(thread_result)}")
                                        thread_emails = []
                                        thread_structure = []
                                
                                # Check if we have any emails from the fetch operations
                                if not thread_emails:
                                    print("No emails retrieved. Trying to use first_message and latest_message if available")
                                    # Try using the first_message and latest_message if available
                                    if isinstance(selected_thread.get('first_message'), dict):
                                        first_msg = selected_thread['first_message']
                                        if 'id' in first_msg and first_msg not in thread_emails:
                                            thread_emails.append(first_msg)
                                            
                                        if isinstance(selected_thread.get('latest_message'), dict):
                                            latest_msg = selected_thread['latest_message']
                                            if 'id' in latest_msg and latest_msg not in thread_emails:
                                                # Check if this is not the same as first_message
                                                if not thread_emails or thread_emails[0].get('id') != latest_msg.get('id'):
                                                    thread_emails.append(latest_msg)
                        
                                # Verify thread emails is a list and contains valid elements
                                if not isinstance(thread_emails, list):
                                    print(f"Error: thread_emails is not a list, it's a {type(thread_emails)}")
                                    thread_emails = []
                                elif len(thread_emails) == 0:
                                    print("Warning: No emails found in thread")
                                else:
                                    print(f"Retrieved {len(thread_emails)} emails for processing")
                                    # Debug print first email structure
                                    print(f"First email keys: {list(thread_emails[0].keys())}")
                    
                            except Exception as e:
                                print(f"Error getting thread emails: {str(e)}")
                                import traceback
                                traceback.print_exc()
                            
                            if thread_emails:
                                # Sort emails by date
                                thread_emails.sort(key=lambda e: e.get('receivedDateTime', ''))
                                print(f"\nProcessing {len(thread_emails)} emails in thread")
                                
                                # Process emails
                                results = []
                                previous_summary = ""
                                
                                progress_bar = st.progress(0)
                                status_text = st.empty()
                                error_placeholder = st.empty()
                                has_rate_limit_error = False
                                
                                for idx, email in enumerate(thread_emails):
                                    progress = (idx + 1) / len(thread_emails)
                                    progress_bar.progress(progress, text=f"Processing email {idx + 1} of {len(thread_emails)}...")
                                    
                                    print(f"\n=== Processing email {idx + 1} ===")
                                    print(f"Subject: {email.get('subject', 'No Subject')}")
                                    print(f"From: {email.get('from', {}).get('emailAddress', {}).get('address', 'Unknown')}")
                                    print(f"Previous summary: {previous_summary[:100]}...")
                                    
                                    try:
                                        # Ensure email has required fields
                                        if not email.get('body', {}).get('content'):
                                            print(f"Warning: Email {idx + 1} has no body content initially. Attempting to fetch full email.")
                                            # Try to get full email content
                                            try:
                                                print(f"Fetching full email content for ID: {email['id']}")
                                                user_email_for_refetch = st.session_state.get('user_email_address')
                                                if not user_email_for_refetch:
                                                    print("Error: User email not found in session state for re-fetching email body.")
                                                else:
                                                    email = graph_client._get_email_with_body(user_email_for_refetch, email['id'])
                                                print(f"Successfully fetched full email content (or attempt made)")
                                            except Exception as e:
                                                print(f"Error fetching full email: {str(e)}")

                                        if email and email.get('body', {}).get('content'):
                                            # Process this individual email with the previous summary
                                            print(f"Processing email {idx + 1} with context from previous emails")
                                            print(f"Previous summary being used: {previous_summary}")
                                            
                                            # Initialize AI service with the Gemini API key
                                            from ai_service import AIService
                                            import google.generativeai as genai

                                            # Fetch Gemini API key from session state
                                            current_gemini_api_key = st.session_state.get("gemini_api_key")

                                            if current_gemini_api_key:
                                                genai.configure(api_key=current_gemini_api_key)
                                                print("DEBUG: Configured genai with API key from session state for thread processing.")
                                            else:
                                                print("DEBUG: Gemini API key not found in session state for thread processing. Genai not configured by this block.")
                                            
                                            ai_service = AIService() # AIService might have its own Gemini key handling
                                            
                                            result = process_individual_email(email, ai_service, previous_summary)
                                            if result:
                                                print(f"Email {idx + 1} processed successfully")
                                                result["email_index"] = idx + 1
                                                results.append(result)
                                                
                                                # Extract summary for next email
                                                if isinstance(result.get("ai_output"), dict) and "Summary" in result["ai_output"]:
                                                    previous_summary = result["ai_output"]["Summary"]
                                                    print(f"Updated previous_summary to: {previous_summary}")
                                                else:
                                                    print("Warning: Could not extract summary from AI output")
                                                
                                                if isinstance(result.get("ai_output"), dict) and result["ai_output"].get("error"):
                                                    has_rate_limit_error = True
                                                    print(f"Warning: Rate limit error in email {idx + 1}")
                                        else:
                                            print(f"Error: Could not get content for email {idx + 1}")
                                        
                                    except Exception as e:
                                        print(f"Error processing email {idx + 1}: {str(e)}")
                                        import traceback
                                        traceback.print_exc()
                                
                                # Complete progress
                                progress_bar.empty()
                                status_text.empty()
                                
                                print("\n=== Results Summary ===")
                                print(f"Total emails processed: {len(results)}")
                                
                                # Store results using our helper function
                                if results:
                                    print("\nAttempting to store results...")
                                    ensure_results_tab_works(results)
                                    
                                    # Verify results were stored
                                    if 'individual_results' in st.session_state and len(st.session_state.individual_results) > 0:
                                        print("Results successfully stored in session state")
                                        
                                        if has_rate_limit_error:
                                            error_placeholder.error("⚠️ Some emails could not be processed due to API rate limits. Results shown may be incomplete.")
                                        else:
                                            st.success(f"Successfully analyzed {len(results)} emails!")
                                            
                                            # Display dashboard and download button
                                            display_results_dashboard()
                                    else:
                                        print("Warning: Results may not have been stored properly")
                                        st.error("Error: Results were not stored properly. Please try again.")
                                else:
                                    st.error("No results to store. This may be due to missing User Email in settings or issues fetching email content.")
                    except Exception as e:
                        st.error(f"An error occurred while processing the thread: {e}")
                        import traceback
                        print("Error details:")
                        print(traceback.format_exc())

# ------------------------------------------------------------
# Tab 2: Detailed Evaluation
# ------------------------------------------------------------

# Helper function to display results dashboard and download button
def display_results_dashboard():
    """Display the metrics dashboard and download button after processing"""
    if not (
        'individual_results' in st.session_state 
        and isinstance(st.session_state.individual_results, list)
        and len(st.session_state.individual_results) > 0
    ):
        return
    
    individual_results = st.session_state.individual_results
    
    # --- ADD DOWNLOAD BUTTON HERE ---
    excel_data = convert_metrics_to_excel(individual_results)
    # Construct a dynamic filename
    thread_subject = get_clean_subject(individual_results[0].get('email', {}).get('subject', 'Thread'))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"evaluation_metrics_{thread_subject[:30].replace(' ', '_')}_{timestamp}.xlsx"
    
    st.download_button(
        label="📥 Download Complete Results Excel",
        data=excel_data,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("---") # Add a separator
    # --- END DOWNLOAD BUTTON ---
    
    # --- START DASHBOARD DISPLAY ---
    st.subheader("Performance Dashboard")
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
        
    # Display overall review 
    st.subheader("Overall Thread Review")
    with st.spinner("Generating overall thread review..."):
        overall_review_text = generate_overall_thread_review(individual_results)
        st.markdown(overall_review_text)
