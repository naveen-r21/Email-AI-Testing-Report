# Email AI Automation

An AI-powered email analysis tool that processes emails from Microsoft Graph API.

## Overview

This application connects to Microsoft Graph API to fetch emails, analyzes them using the Dwellworks AI API, and displays the results.

## How to Fix app_backup.py

If you encounter syntax errors in `app_backup.py`, follow these steps:

1. Use the fixed function from `email_sequential_processor.py` instead:

```python
# Import the fixed function
from email_sequential_processor import process_thread_emails

# Use this instead of the broken function in app_backup.py
```

2. Alternative fix - use `app_working.py` as your main application:

```bash
# Rename files to use the working version
mv app_backup.py app_backup_old.py
mv app_working.py app.py
```

3. You can also copy specific fixed functions from `app.py` to `app_backup.py` to fix it:

```python
def process_thread_emails(thread_emails):
    """Process emails through the API sequentially with summary chaining"""
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
            
            # Evaluate results
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
```

## Installation

1. Clone the repository
2. Install the required packages: `pip install -r requirements.txt`
3. Run the application: `streamlit run app.py`

## Features

- Fetch emails from Microsoft Graph API
- Process and analyze emails using AI
- Display results with metrics
- Download analysis reports

## Email Threading

The application uses an improved method for organizing email threads based on the relationship between messages:

### Thread Organization

- Emails are linked using `internetMessageId` and `inReplyTo` fields
- Each message has a globally unique ID (`internetMessageId`)
- Reply messages contain a reference to their parent message (`inReplyTo`)
- This builds a proper hierarchical thread structure

### Benefits

- More accurate thread reconstruction compared to using just conversation IDs
- Works for large threads (50-100+ emails) with O(n) complexity
- Clear visualization of thread hierarchy in the UI
- Resilient to various email client formatting differences

The thread structure is visualized in the main interface, showing the parent-child relationships between messages with proper indentation.

## Email Threading Improvements

### Subject-Based Thread Grouping

The application now includes an improved method for organizing email threads based on subject similarity:

- Groups emails with similar subjects together, even when conversation IDs differ
- Uses fuzzy matching to handle minor variations in subject lines (85% similarity threshold)
- Properly handles "Re:" prefixed replies as part of the same thread
- More intuitive thread organization that matches how users think about emails

### Metadata Cleaning

Email content is now cleaned more effectively:

- Removes metadata and signatures by cutting content after "from:" appears
- Preserves important information like property details, dates, and times
- Improves AI analysis by focusing on the actual email message
- Removes disclaimers, forwarded content, and other noise

These improvements make the application more reliable when working with emails from different sources or email clients that may handle threading differently.

## License

This project is licensed under the MIT License.