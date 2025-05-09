"""
Vector store for email data storage and retrieval.
Integrates with your existing solution to store and query emails.
"""

import os
import json
import pandas as pd
from datetime import datetime
from typing import List, Dict, Any, Optional

class VectorStore:
    """Simple file-based vector store for emails"""
    
    def __init__(self, storage_dir: str = "email_data"):
        self.storage_dir = storage_dir
        os.makedirs(storage_dir, exist_ok=True)
    
    def store_emails(self, emails: List[Dict[str, Any]], conversation_id: str) -> str:
        """Store a batch of emails with the same conversation ID"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{self.storage_dir}/conversation_{conversation_id}_{timestamp}.json"
        
        with open(filename, 'w') as f:
            json.dump(emails, f, indent=2)
        
        return filename
    
    def get_all_conversations(self) -> List[str]:
        """Get list of all conversation IDs from stored data"""
        conversations = set()
        for filename in os.listdir(self.storage_dir):
            if filename.startswith("conversation_") and filename.endswith(".json"):
                # Extract conversation ID from filename
                parts = filename.split("_")
                if len(parts) > 1:
                    conversations.add(parts[1])
        
        return list(conversations)
    
    def get_emails_by_conversation(self, conversation_id: str) -> List[Dict[str, Any]]:
        """Get all emails for a specific conversation ID"""
        emails = []
        for filename in os.listdir(self.storage_dir):
            if f"conversation_{conversation_id}_" in filename and filename.endswith(".json"):
                with open(os.path.join(self.storage_dir, filename), 'r') as f:
                    emails.extend(json.load(f))
        
        # Sort by date
        emails.sort(key=lambda x: x.get('receivedDateTime', ''))
        return emails
    
    def save_results(self, results: Dict[str, Any], result_type: str = "evaluation") -> str:
        """Save evaluation results to disk"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{self.storage_dir}/{result_type}_{timestamp}.json"
        
        with open(filename, 'w') as f:
            json.dump(results, f, indent=2)
        
        return filename
    
    def export_to_excel(self, data: Dict[str, Any], filename: Optional[str] = None) -> str:
        """Export results to Excel with the specified format"""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{self.storage_dir}/report_{timestamp}.xlsx"
        
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            # Create Input sheet
            if 'input' in data:
                df_input = pd.DataFrame(data['input'])
                df_input.to_excel(writer, sheet_name='Input', index=False)
            
            # Create Groundtruth sheet
            if 'groundtruth' in data:
                df_groundtruth = pd.DataFrame(data['groundtruth'])
                df_groundtruth.to_excel(writer, sheet_name='Groundtruth', index=False)
            
            # Create Output sheet
            if 'output' in data:
                df_output = pd.DataFrame(data['output'])
                df_output.to_excel(writer, sheet_name='Output', index=False)
            
            # Create Metrics sheet if available
            if 'metrics' in data:
                df_metrics = pd.DataFrame([data['metrics']])
                df_metrics.to_excel(writer, sheet_name='Metrics', index=False)
        
        return filename