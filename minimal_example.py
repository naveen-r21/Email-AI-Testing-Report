import streamlit as st
import pandas as pd
from io import BytesIO

# Set page title
st.set_page_config(page_title="Excel Email Uploader", layout="wide")

# Set up session state if needed
if "email_source" not in st.session_state:
    st.session_state.email_source = "Upload emails as Excel"

# Display title
st.title("Email Excel Uploader")

# Create the radio selection
selected_option = st.radio(
    "Select Email Source:", 
    options=["Upload emails as Excel", "Fetch emails from Outlook"],
    index=0 if st.session_state.email_source == "Upload emails as Excel" else 1,
    horizontal=True
)

# Update session state
st.session_state.email_source = selected_option

# Display debug info
st.write(f"Current selection: {selected_option}")

# Show Excel upload UI when selected
if selected_option == "Upload emails as Excel":
    st.subheader("Upload Excel File")
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Upload an Excel file with an 'Email Content' column",
        type=["xlsx", "xls"]
    )
    
    if uploaded_file is not None:
        # Try to read and display preview
        try:
            df = pd.read_excel(uploaded_file)
            st.success("File uploaded successfully!")
            
            st.subheader("File Preview")
            st.dataframe(df.head())
            
            # Check for required column
            if "Email Content" in df.columns:
                st.success("✅ 'Email Content' column found!")
                
                # Process button
                if st.button("Process Emails", type="primary"):
                    with st.spinner("Processing emails..."):
                        # Simulate processing
                        st.session_state.processed = True
                        st.success(f"Processed {len(df)} emails!")
            else:
                st.error("❌ 'Email Content' column not found. Please check your file.")
                st.write("Available columns:", ", ".join(df.columns.tolist()))
                
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
# Show Outlook UI when selected
else:
    st.subheader("Outlook Email Processing")
    st.info("This section would contain Outlook fetching functionality.") 