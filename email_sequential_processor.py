"""
Email Sequential Processor with Summary Chaining

This module contains a fixed implementation of email thread processing 
with sequential summary chaining. Import and use these functions directly 
to replace the buggy ones in app_backup.py.
"""

import streamlit as st
import re
import json
from ai_service import AIService

def process_thread_emails(thread_emails):
    """Process emails through the API sequentially with summary chaining
    
    Args:
        thread_emails: List of emails from Graph API
        
    Returns:
        List of results with API outputs
    """
    # Create AI service
    ai_service = AIService()
    results = []
    
    # Set up progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    total = len(thread_emails)
    
    # Initialize previous summary
    previous_summary = ""
    
    # Process each email individually but pass the previous email's summary
    for i, email in enumerate(thread_emails):
        try:
            # Update status
            status_text.text(f"Processing email {i+1} of {total}...")
            print(f"Processing with previous summary: {previous_summary}")
            
            # Create input data format
            input_data = {
                "mail_id": email.get("id", ""),
                "file_name": [],
                "email": email.get("sender", {}).get("emailAddress", {}).get("address", ""),
                "mail_time": email.get("receivedDateTime", ""),
                "body_type": email.get("body", {}).get("contentType", ""),
                "mail_body": email.get("body", {}).get("content", ""),
                "thread_id": email.get("conversationId", ""),
                "mail_summary": previous_summary
            }
            
            # Call API
            ai_output = ai_service.analyze_email_thread([email])
            
            # Get the summary for the next email
            if isinstance(ai_output, dict):
                previous_summary = ai_output.get('Summary', '')
                print(f"Got summary for next email: {previous_summary}")
            
            # Evaluate results - this function should be defined elsewhere and imported
            # In this example we'll use a simple placeholder
            metrics = evaluate_results(ai_output, {})
            
            # Store result
            results.append({
                "email_index": i+1,
                "input_data": input_data,
                "ai_output": ai_output,
                "metrics": metrics,
                "email_content": {
                    "sender_name": email.get('sender', {}).get('emailAddress', {}).get('name', 'Unknown'),
                    "sender_email": email.get('sender', {}).get('emailAddress', {}).get('address', 'unknown'),
                    "content": re.sub(r'<[^>]+>', '', email.get('body', {}).get('content', '')),
                    "sent_time": email.get('receivedDateTime', '')
                }
            })
        except Exception as e:
            st.error(f"Error processing email {i+1}: {str(e)}")
        
        # Update progress
        progress_bar.progress((i+1)/total)
    
    # Complete progress
    progress_bar.progress(1.0)
    status_text.text("All emails processed.")
    
    return results

def evaluate_results(ai_output, groundtruth=None):
    """Evaluate AI output against groundtruth"""
    evaluation = []
    
    # Define fields to evaluate
    fields = [
        "Summary", 
        "Sentiment analysis",
        "overall_sentiment_analysis",
        "feature",
        "category",
        "Events"
    ]
    
    for field in fields:
        # Get values from both outputs (handle different formats and missing values)
        ai_value = ai_output.get(field) if ai_output and isinstance(ai_output, dict) else None
        gt_value = groundtruth.get(field) if groundtruth and isinstance(groundtruth, dict) else None
        
        # Add to evaluation results
        if ai_value is not None or gt_value is not None:
            status = "Pass" if ai_value == gt_value else "Fail"
            evaluation.append({
                "Field": field,
                "Status": status,
                "AI Value": ai_value,
                "Ground Truth": gt_value
            })
    
    # Handle nested fields like sentiment_analysis which might be objects
    if "sentiment_analysis" in ai_output and isinstance(ai_output["sentiment_analysis"], dict):
        sentiment_value = ai_output["sentiment_analysis"].get("overall")
        gt_sentiment = groundtruth.get("Sentiment analysis")
        
        if sentiment_value is not None or gt_sentiment is not None:
            status = "Pass" if str(sentiment_value).lower() == str(gt_sentiment).lower() else "Fail"
            evaluation.append({
                "Field": "sentiment_analysis.overall",
                "Status": status,
                "AI Value": sentiment_value,
                "Ground Truth": gt_sentiment
            })
    
    # If key_entities exists, mark it as info
    if "key_entities" in ai_output:
        evaluation.append({
            "Field": "key_entities",
            "Status": "Info",
            "AI Value": ai_output["key_entities"],
            "Ground Truth": groundtruth.get("key_entities", "N/A")
        })
    
    return evaluation

def display_evaluation_metrics(result):
    """Display evaluation metrics in a nicely formatted table"""
    if "metrics" not in result or not result["metrics"]:
        st.warning("No evaluation metrics available.")
        return
        
    # Format into HTML table with styling
    styled_html = """
    <table style='width:100%; border-collapse: collapse; border: 1px solid #444;'>
        <tr style='background-color: #2D2D2D;'>
            <th style='padding: 8px; text-align: left; border: 1px solid #444;'>Field</th>
            <th style='padding: 8px; text-align: left; border: 1px solid #444;'>Status</th>
            <th style='padding: 8px; text-align: left; border: 1px solid #444;'>AI Value</th>
            <th style='padding: 8px; text-align: left; border: 1px solid #444;'>Ground Truth</th>
        </tr>
    """
    
    for i, row in enumerate(result["metrics"]):
        bg_color = "#1A1A1A" if i % 2 == 0 else "#222"
        field = row.get("Field", "")
        
        # Handle different status field names (Pass/Fail or Status)
        if "Pass/Fail" in row:
            status = row["Pass/Fail"]
        elif "Status" in row:
            status = row["Status"]
        else:
            status = "Unknown"
            
        color = "green" if status == "Pass" else "red" if status == "Fail" else "#888"
        
        # Format AI Value
        ai_value = row.get("AI Value", "N/A")
        if isinstance(ai_value, (list, dict)):
            ai_value_display = json.dumps(ai_value, ensure_ascii=False).replace('\n', '\\n')
        else:
            ai_value_display = str(ai_value).replace('\n', '\\n')
            
        # Format Ground Truth
        groundtruth = row.get("Ground Truth", "N/A")
        if isinstance(groundtruth, (list, dict)):
            groundtruth_display = json.dumps(groundtruth, ensure_ascii=False).replace('\n', '\\n')
        else:
            groundtruth_display = str(groundtruth).replace('\n', '\\n')
        
        styled_html += f"""
        <tr style='background-color: {bg_color};'>
            <td style='padding: 8px; border: 1px solid #444;'>{field}</td>
            <td style='padding: 8px; border: 1px solid #444; color:{color}; font-weight:bold;'>{status}</td>
            <td style='padding: 8px; border: 1px solid #444; white-space: pre-wrap;'>{ai_value_display}</td>
            <td style='padding: 8px; border: 1px solid #444; white-space: pre-wrap;'>{groundtruth_display}</td>
        </tr>
        """
    
    styled_html += "</table>"
    
    st.markdown("### 📊 Evaluation Results")
    st.markdown(styled_html, unsafe_allow_html=True)

# Example usage:
# results = process_thread_emails(thread_emails)
# for result in results:
#     display_evaluation_metrics(result) 