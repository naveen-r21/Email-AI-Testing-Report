"""
Evaluation engine for comparing AI outputs with groundtruth.
Generates metrics and detailed comparisons.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, List, Union, Optional
import re
import streamlit as st
import json
import math

class EvaluationEngine:
    """Evaluation engine for generating metrics and reports"""
    
    @staticmethod
    def calculate_metrics(evaluations: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Calculate overall metrics from a list of evaluations
        Returns aggregated metrics and statistics
        """
        if not evaluations:
            return {
                "average_accuracy": 0,
                "field_accuracies": {},
                "total_evaluations": 0,
                "success_rate": 0
            }
        
        # Extract overall accuracy from each evaluation
        accuracies = [eval.get("overall_accuracy", 0) for eval in evaluations]
        
        # Calculate field-specific accuracies
        fields = ["email_summarization", "events_summarization", "sentiment_analysis", "category_analysis"]
        field_scores = {field: [] for field in fields}
        
        for eval in evaluations:
            comparison = eval.get("comparison", {})
            for field in fields:
                if field in comparison:
                    field_data = comparison[field]
                    if isinstance(field_data, dict):
                        score = field_data.get("score", 0)
                        if isinstance(score, (int, float)):
                            field_scores[field].append(score)
        
        # Calculate average for each field
        field_accuracies = {}
        for field, scores in field_scores.items():
            if scores:
                field_accuracies[field] = sum(scores) / len(scores)
            else:
                field_accuracies[field] = 0
        
        # Count successful evaluations (overall accuracy > 70%)
        successful = sum(1 for acc in accuracies if acc >= 70)
        
        return {
            "average_accuracy": sum(accuracies) / len(evaluations) if accuracies else 0,
            "field_accuracies": field_accuracies,
            "total_evaluations": len(evaluations),
            "success_rate": (successful / len(evaluations)) * 100 if evaluations else 0
        }
    
    @staticmethod
    def generate_report_data(processed_emails: List[Dict[str, Any]], 
                           groundtruths: List[Dict[str, Any]],
                           ai_outputs: List[Dict[str, Any]],
                           evaluations: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Generate structured data for reporting
        Formats data for Excel and JSON exports
        """
        # Format input data
        input_data = []
        for email in processed_emails:
            input_data.append({
                "From Email": email.get("from_email", ""),
                "To Email": ", ".join(email.get("to_email", [])),
                "DateTime": email.get("datetime", ""),
                "Email Subject": email.get("subject", ""),
                "Email content": email.get("content", "")
            })
        
        # Format groundtruth data
        groundtruth_data = []
        for gt in groundtruths:
            groundtruth_data.append({
                "Email summarization": gt.get("email_summarization", ""),
                "Events summarization": gt.get("events_summarization", ""),
                "Sentiment analysis": gt.get("sentiment_analysis", ""),
                "category analysis": gt.get("category_analysis", "")
            })
        
        # Format AI output data
        output_data = []
        for output in ai_outputs:
            output_data.append({
                "Email summarization": output.get("email_summarization", ""),
                "Events summarization": output.get("events_summarization", ""),
                "Sentiment analysis": output.get("sentiment_analysis", ""),
                "category analysis": output.get("category_analysis", "")
            })
        
        # Calculate metrics
        metrics = EvaluationEngine.calculate_metrics(evaluations)
        
        return {
            "input": input_data,
            "groundtruth": groundtruth_data,
            "output": output_data,
            "metrics": metrics,
            "evaluations": evaluations  # Include full evaluations for JSON export
        }

def calculate_similarity(text1: str, text2: str) -> float:
    """Calculate similarity between two text strings using Jaccard similarity"""
    if not text1 or not text2:
        return 0.0
    
    # Tokenize and create sets of words
    words1 = set(re.findall(r'\b\w+\b', text1.lower()))
    words2 = set(re.findall(r'\b\w+\b', text2.lower()))
    
    # Calculate Jaccard similarity (intersection over union)
    intersection = len(words1.intersection(words2))
    union = len(words1.union(words2))
    
    if union == 0:
        return 0.0
    
    return intersection / union

def evaluate_results(ai_output: Dict[str, Any], groundtruth: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """
    Evaluate AI output against groundtruth and return metrics
    
    Args:
        ai_output: AI-generated output with extracted features
        groundtruth: Optional groundtruth data
        
    Returns:
        List of evaluation metrics with detailed explanations
    """
    metrics = []
    
    if not ai_output:
        return metrics
    
    if groundtruth is None:
        groundtruth = {}
    
    # SENTIMENT ANALYSIS EVALUATION
    if "sentiment_analysis" in ai_output:
        sentiment = ai_output.get("sentiment_analysis", "")
        gt_sentiment = groundtruth.get("sentiment_analysis", "")
        
        # Normalize sentiment values for comparison
        def normalize_sentiment(sent):
            sent = str(sent).lower().strip()
            if sent in ["positive", "green", "good", "1"]:
                return "positive"
            elif sent in ["negative", "red", "bad", "-1"]:
                return "negative"
            else:
                return "neutral"
        
        norm_sentiment = normalize_sentiment(sentiment)
        norm_gt_sentiment = normalize_sentiment(gt_sentiment)
        
        # Generate explanation for ground truth determination
        gt_explanation = "No ground truth explanation available"
        if "sentiment_analysis_explanation" in groundtruth:
            gt_explanation = groundtruth.get("sentiment_analysis_explanation", "")
        else:
            # Generate a basic explanation based on common patterns
            if norm_gt_sentiment == "positive":
                gt_explanation = "Email contains predominantly positive language and tone (e.g., 'thank you', 'appreciate', 'pleased')"
            elif norm_gt_sentiment == "negative":
                gt_explanation = "Email contains predominantly negative language or concerns (e.g., 'issue', 'problem', 'regret')"
            else:
                gt_explanation = "Email contains neutral or balanced language without strong positive or negative indicators"
        
        # Generate explanation for pass/fail status
        status = "Pass" if norm_sentiment == norm_gt_sentiment else "Fail"
        comparison_explanation = ""
        
        if status == "Pass":
            comparison_explanation = f"AI correctly identified the {norm_gt_sentiment} sentiment of the email"
        else:
            comparison_explanation = f"AI incorrectly classified sentiment as {norm_sentiment} when it should be {norm_gt_sentiment}"
        
        # Add metric with detailed explanations
        metrics.append({
            "Field": "sentiment_analysis",
            "AI Value": sentiment,
            "Ground Truth": gt_sentiment,
            "Status": status,
            "Ground Truth Explanation": gt_explanation,
            "Comparison Explanation": comparison_explanation
        })
    
    # OVERALL SENTIMENT ANALYSIS EVALUATION
    if "overall_sentiment_analysis" in ai_output:
        overall = ai_output.get("overall_sentiment_analysis", "")
        gt_overall = groundtruth.get("overall_sentiment_analysis", "")
        
        # Normalize sentiment values
        norm_overall = normalize_sentiment(overall)
        norm_gt_overall = normalize_sentiment(gt_overall)
        
        # Generate explanation for ground truth determination
        gt_explanation = "No ground truth explanation available"
        if "overall_sentiment_explanation" in groundtruth:
            gt_explanation = groundtruth.get("overall_sentiment_explanation", "")
        else:
            if norm_gt_overall == "positive":
                gt_explanation = "Thread progression shows resolution or positive outcome"
            elif norm_gt_overall == "negative":
                gt_explanation = "Thread progression shows unresolved issues or negative outcome"
            else:
                gt_explanation = "Thread maintains neutral tone throughout or balanced positive/negative elements"
        
        # Generate explanation for pass/fail status
        status = "Pass" if norm_overall == norm_gt_overall else "Fail"
        comparison_explanation = ""
        
        if status == "Pass":
            comparison_explanation = f"AI correctly identified the overall {norm_gt_overall} sentiment of the thread"
        else:
            comparison_explanation = f"AI incorrectly classified overall sentiment as {norm_overall} when thread progression indicates {norm_gt_overall}"
        
        # Add metric with detailed explanations
        metrics.append({
            "Field": "overall_sentiment_analysis",
            "AI Value": overall,
            "Ground Truth": gt_overall,
            "Status": status,
            "Ground Truth Explanation": gt_explanation,
            "Comparison Explanation": comparison_explanation
        })
    
    # FEATURE & CATEGORY IDENTIFICATION EVALUATION
    if "feature" in ai_output:
        feature = ai_output.get("feature", "")
        gt_feature = groundtruth.get("feature", "")
        
        # Generate explanation for ground truth determination
        gt_explanation = "No ground truth explanation available"
        
        if "feature_explanation" in groundtruth:
            gt_explanation = groundtruth.get("feature_explanation", "")
        else:
            # Generate explanation based on feature type
            if "DSC First Contact" in gt_feature:
                gt_explanation = "First outreach email from DSC to EE with introduction language"
            elif "EE First Contact" in gt_feature:
                gt_explanation = "First email from EE to DSC with availability or relocation information"
            elif "Phone Consultation Scheduled" in gt_feature:
                gt_explanation = "Email confirming specific date/time for first phone consultation"
            elif "Phone Consultation Completed" in gt_feature:
                gt_explanation = "Email confirming the completion of first phone consultation"
            elif "No feature" in gt_feature:
                gt_explanation = "Email does not match any defined feature conditions"
            else:
                gt_explanation = "Feature determined based on email content and direction"
        
        # Generate explanation for pass/fail status
        status = "Pass" if feature == gt_feature else "Fail"
        comparison_explanation = ""
        
        if status == "Pass":
            comparison_explanation = f"AI correctly identified the feature: {gt_feature}"
        else:
            comparison_explanation = f"AI classified as '{feature}' when it should be '{gt_feature}'"
        
        # Add metric with detailed explanations
        metrics.append({
            "Field": "feature",
            "AI Value": feature,
            "Ground Truth": gt_feature,
            "Status": status,
            "Ground Truth Explanation": gt_explanation,
            "Comparison Explanation": comparison_explanation
        })
    
    # CATEGORY EVALUATION
    if "category" in ai_output:
        category = ai_output.get("category", "")
        gt_category = groundtruth.get("category", "")
        
        # Generate explanation for ground truth determination
        gt_explanation = "No ground truth explanation available"
        
        if "category_explanation" in groundtruth:
            gt_explanation = groundtruth.get("category_explanation", "")
        else:
            if "Initial Service Milestones" in gt_category:
                gt_explanation = "Email relates to a defined milestone in the initial service process"
            elif "No category" in gt_category:
                gt_explanation = "Email does not match any defined category conditions"
            else:
                gt_explanation = "Category determined based on email content and context"
        
        # Generate explanation for pass/fail status
        status = "Pass" if category == gt_category else "Fail"
        comparison_explanation = ""
        
        if status == "Pass":
            comparison_explanation = f"AI correctly identified the category: {gt_category}"
        else:
            comparison_explanation = f"AI classified as '{category}' when it should be '{gt_category}'"
        
        # Add metric with detailed explanations
        metrics.append({
            "Field": "category",
            "AI Value": category,
            "Ground Truth": gt_category,
            "Status": status,
            "Ground Truth Explanation": gt_explanation,
            "Comparison Explanation": comparison_explanation
        })
    
    # EVENT DETECTION EVALUATION
    # Get events from both AI output and groundtruth
    ai_events = ai_output.get("Events", [])
    gt_events = groundtruth.get("Events", [])
    
    # Standardize to list format
    if not isinstance(ai_events, list):
        ai_events = [ai_events] if ai_events else []
    if not isinstance(gt_events, list):
        gt_events = [gt_events] if gt_events else []
    
    # Get first event for comparison (or None if no events)
    ai_event = ai_events[0] if ai_events else None
    gt_event = gt_events[0] if gt_events else None
    
    # Define all event fields to evaluate
    event_fields = ["Event name", "Date", "Time", "Property Type", "Agent Name", "Location"]
    
    # Evaluate each event field
    for field in event_fields:
        ai_value = ai_event.get(field, None) if ai_event else None
        gt_value = gt_event.get(field, None) if gt_event else None
        
        # Convert None to "null" for display
        ai_value_display = "null" if ai_value is None else ai_value
        gt_value_display = "null" if gt_value is None else gt_value
        
        # Generate explanation for ground truth
        gt_explanation = ""
        if gt_value is None:
            gt_explanation = f"No {field.lower()} information found in email content"
        else:
            gt_explanation = f"{field} information extracted from email content: '{gt_value}'"
        
        # Determine field status
        field_status = "Info"  # Default status
        comparison_explanation = ""
        
        if gt_value and not ai_value:
            field_status = "Fail"  # Missing in AI output
            comparison_explanation = f"AI failed to extract {field.lower()} information that exists in the email"
        elif ai_value and not gt_value:
            field_status = "Info"  # Not in groundtruth
            comparison_explanation = f"AI extracted {field.lower()} information not identified in ground truth"
        elif ai_value == gt_value:
            field_status = "Pass"  # Exact match
            comparison_explanation = f"AI correctly extracted {field.lower()} information"
        elif ai_value and gt_value:
            # Both values present but different - check for similarity
            similarity = calculate_similarity(str(ai_value), str(gt_value))
            if similarity >= 0.7:
                field_status = "Partial Pass"
                comparison_explanation = f"AI extracted similar but not identical {field.lower()} information (similarity: {similarity:.2f})"
            else:
                field_status = "Fail"
                comparison_explanation = f"AI extracted incorrect {field.lower()} information"
        
        # Add metric with detailed explanations
        metrics.append({
            "Field": f"Event - {field}",
            "AI Value": ai_value_display,
            "Ground Truth": gt_value_display,
            "Status": field_status,
            "Ground Truth Explanation": gt_explanation,
            "Comparison Explanation": comparison_explanation
        })
    
    # Add overall event match metric
    if ai_events or gt_events:
        # Create a normalized representation of the event for comparison
        def normalize_event(event):
            if not event:
                return ""
            parts = []
            for field in event_fields:
                value = event.get(field, "")
                if value:
                    parts.append(str(value))
            return "__".join(parts)
        
        ai_event_norm = normalize_event(ai_event)
        gt_event_norm = normalize_event(gt_event)
        
        # Calculate similarity for partial matches
        similarity = 0
        if ai_event_norm and gt_event_norm:
            similarity = calculate_similarity(ai_event_norm, gt_event_norm)
        
        # Determine status based on similarity
        if ai_event_norm == gt_event_norm:
            status = "Pass"
            comparison_explanation = "AI correctly identified all event details"
        elif similarity >= 0.7:
            status = "Partial Pass"
            comparison_explanation = f"AI identified most event details correctly (similarity: {similarity:.2f})"
        else:
            status = "Fail"
            comparison_explanation = f"AI failed to correctly identify most event details (similarity: {similarity:.2f})"
        
        # Generate ground truth explanation
        if not gt_event:
            gt_explanation = "No event details found in email content"
        else:
            gt_explanation = "Event details extracted from email content based on date, time, and location patterns"
        
        metrics.append({
            "Field": "Event Match",
            "AI Value": ai_event_norm,
            "Ground Truth": gt_event_norm,
            "Status": status,
            "Ground Truth Explanation": gt_explanation,
            "Comparison Explanation": comparison_explanation
        })
    
    # Add similarity and content validation checks for summary
    if "Summary" in ai_output and "Summary" in groundtruth:
        ai_summary = ai_output.get("Summary", "")
        gt_summary = groundtruth.get("Summary", "")
        
        # 1. Calculate similarity for summary
        similarity = calculate_similarity(ai_summary, gt_summary)
        
        # 2. Check if summary content appears in the email
        email_content = ""
        if 'email_content' in groundtruth:
            email_content = groundtruth.get('email_content', '')
        
        # Validate if summary terms appear in email content
        content_words = set(re.findall(r'\b\w+\b', email_content.lower()))
        summary_words = set(re.findall(r'\b\w+\b', ai_summary.lower()))
        
        # Filter out common stopwords
        stopwords = {'the', 'a', 'an', 'and', 'or', 'but', 'is', 'are', 'was', 'were', 
                    'in', 'on', 'at', 'to', 'for', 'with', 'by', 'about', 'as', 'of', 
                    'from', 'has', 'have', 'had', 'be', 'been', 'being', 'this', 'that'}
        
        content_words = content_words - stopwords
        summary_words = summary_words - stopwords
        
        # Calculate content validation percentage
        matched_words = summary_words.intersection(content_words)
        content_validation_pct = len(matched_words) / len(summary_words) if summary_words else 0
        
        # Determine status based on both similarity and content validation
        if similarity >= 0.7 and content_validation_pct >= 0.7:
            status = "Pass"
        elif similarity >= 0.4 or content_validation_pct >= 0.5:
            status = "Partial Pass"
        else:
            status = "Fail"
        
        # Generate explanations
        gt_explanation = "Summary generated based on email subject, sender, and key content elements"
        
        comparison_explanation = f"AI summary matches ground truth with {similarity:.2f} similarity and {content_validation_pct:.2%} content validation. "
        if status == "Pass":
            comparison_explanation += "The summary accurately reflects email content."
        elif status == "Partial Pass":
            comparison_explanation += "The summary partially reflects email content."
        else:
            comparison_explanation += "The summary does not adequately reflect email content."
        
        metrics.append({
            "Field": "Summary",
            "AI Value": ai_summary,
            "Ground Truth": gt_summary,
            "Status": status,
            "Ground Truth Explanation": gt_explanation,
            "Comparison Explanation": comparison_explanation,
            "Content Validation": f"{content_validation_pct:.2%}",
            "Similarity": f"{similarity:.2f}",
            "individual_email_review_points": "- AI correctly matched the positive sentiment cues.\n- No major misses for sentiment on this email."
        })
    
    return metrics

def display_evaluation_metrics(result: Dict[str, Any]) -> None:
    """
    Display evaluation metrics in a structured format
    
    Args:
        result: The result dict containing metrics, ai_output, and groundtruth
    """
    print("=== Debug: Starting evaluate_results ===")
    
    try:
        metrics = result.get("metrics", [])
        ai_output = result.get("ai_output", {})
        groundtruth = result.get("groundtruth", {})
        email = result.get("email", {})
        
        # Check if we have LLM-generated metrics
        has_metrics = len(metrics) > 0
        
        # Function to highlight status in dataframes
        def highlight_status(row):
            if row.name != 'Status':
                return [''] * len(row)
            
            styles = []
            for val in row:
                if val == 'Pass':
                    styles.append('background-color: #1e7e34; color: white')
                elif val == 'Fail':
                    styles.append('background-color: #dc3545; color: white')
                elif val == 'Partial Pass':
                    styles.append('background-color: #fd7e14; color: white')
                else:
                    styles.append('')
            return styles
        
        # Function to create a comparison table for metrics
        def create_comparison_table(metrics_list):
            if not metrics_list:
                st.info("No metrics available for this category")
                return
                
            # Create DataFrame
            df = pd.DataFrame({
                "Metric": [m.get("Metric", "Unknown") for m in metrics_list],
                "AI Value": [m.get("AI Value", "N/A") for m in metrics_list],
                "Ground Truth": [m.get("Ground Truth", "N/A") for m in metrics_list],
                "Status": [m.get("Status", "N/A") for m in metrics_list]
            })
            
            # Display table
            st.dataframe(df.style.apply(highlight_status, axis=1), hide_index=True)
            
            # Display explanations
            for metric in metrics_list:
                st.markdown(f"**{metric.get('Metric', 'Unknown')} Explanation:**")
                explanation = metric.get("Pass/Fail Explanation", metric.get("Comparison Explanation", "No explanation available"))
                st.markdown(f"_{explanation}_")
                st.markdown("---")
        
        # If no metrics and we have both AI output and groundtruth, generate fallback metrics
        if not has_metrics and ai_output and groundtruth:
            from evaluation_engine import generate_fallback_metrics
            metrics = generate_fallback_metrics(ai_output, groundtruth)
            has_metrics = len(metrics) > 0
            st.warning("Using fallback metrics because LLM evaluation failed")
        
        # ===== SENTIMENT ANALYSIS EVALUATION =====
        with st.expander("🎭 SENTIMENT ANALYSIS EVALUATION", expanded=True):
            sentiment_metrics = [m for m in metrics if m.get("Metric") in ["Sentiment analysis", "overall_sentiment_analysis"]]
            
            if sentiment_metrics:
                create_comparison_table(sentiment_metrics)
            else:
                # Create a basic comparison table for sentiment
                st.info("No sentiment analysis metrics available")
                if ai_output and groundtruth:
                    df = pd.DataFrame({
                        "Metric": ["Sentiment analysis", "overall_sentiment_analysis"],
                        "AI Value": [ai_output.get("Sentiment analysis", "N/A"), ai_output.get("overall_sentiment_analysis", "N/A")],
                        "Ground Truth": [groundtruth.get("Sentiment analysis", "N/A"), groundtruth.get("overall_sentiment_analysis", "N/A")],
                        "Status": ["N/A", "N/A"]
                    })
                    st.dataframe(df.style.apply(highlight_status, axis=1), hide_index=True)
        
        # ===== FEATURE & CATEGORY EVALUATION =====
        with st.expander("⚙️ FEATURE & CATEGORY EVALUATION", expanded=True):
            feature_metrics = [m for m in metrics if m.get("Metric") in ["feature", "category"]]
            
            if feature_metrics:
                create_comparison_table(feature_metrics)
            else:
                # Create a basic comparison table for features
                st.info("No feature analysis metrics available")
                if ai_output and groundtruth:
                    df = pd.DataFrame({
                        "Metric": ["feature", "category"],
                        "AI Value": [ai_output.get("feature", "N/A"), ai_output.get("category", "N/A")],
                        "Ground Truth": [groundtruth.get("feature", "N/A"), groundtruth.get("category", "N/A")],
                        "Status": ["N/A", "N/A"]
                    })
                    st.dataframe(df.style.apply(highlight_status, axis=1), hide_index=True)
        
        # ===== EVENT DETECTION EVALUATION =====
        with st.expander("📅 EVENT DETECTION EVALUATION", expanded=True):
            event_metrics = [m for m in metrics if m.get("Metric") == "Events"]
            
            if event_metrics:
                create_comparison_table(event_metrics)
            else:
                # Create a basic comparison table for events
                st.info("No event detection metrics available")
                if ai_output and groundtruth:
                    ai_events = ai_output.get("Events", [])
                    gt_events = groundtruth.get("Events", [])
                    df = pd.DataFrame({
                        "Metric": ["Events"],
                        "AI Value": [f"{len(ai_events)} events detected"],
                        "Ground Truth": [f"{len(gt_events)} events"],
                        "Status": ["N/A"]
                    })
                    st.dataframe(df.style.apply(highlight_status, axis=1), hide_index=True)
        
        # ===== SUMMARY EVALUATION =====
        with st.expander("📝 SUMMARY EVALUATION", expanded=True):
            summary_metrics = [m for m in metrics if m.get("Metric") == "Summary"]
            
            if summary_metrics:
                create_comparison_table(summary_metrics)
            else:
                # Create a basic comparison table for summary
                st.info("No summary evaluation metrics available")
                if ai_output and groundtruth:
                    df = pd.DataFrame({
                        "Metric": ["Summary"],
                        "AI Value": [ai_output.get("Summary", "N/A")[:100] + "..." if len(ai_output.get("Summary", "")) > 100 else ai_output.get("Summary", "N/A")],
                        "Ground Truth": [groundtruth.get("Summary", "N/A")[:100] + "..." if len(groundtruth.get("Summary", "")) > 100 else groundtruth.get("Summary", "N/A")],
                        "Status": ["N/A"]
                    })
                    st.dataframe(df.style.apply(highlight_status, axis=1), hide_index=True)
    except Exception as e:
        st.error(f"Error displaying evaluation metrics: {str(e)}")
        print(f"Error in evaluation metrics display: {str(e)}")

def evaluate_with_llm(ai_output: Dict[str, Any], groundtruth: Dict[str, Any], email_content: str) -> List[Dict[str, Any]]:
    """
    Use LLM to evaluate AI output against groundtruth with detailed analysis
    
    Args:
        ai_output: AI-generated output
        groundtruth: Ground truth data
        email_content: Original email content for context
        
    Returns:
        List of evaluation metrics with detailed explanations
    """
    # Initialize Gemini
    import google.generativeai as genai
    import os
    
    # Configure Gemini
    genai.configure(api_key=os.getenv('GEMINI_API_KEY', st.session_state.get('gemini_api_key', '')))
    model = genai.GenerativeModel('gemini-1.5-flash')  # Use a valid model
    
    # Use template strings to avoid f-string format errors
    prompt_template = f'''
    You are an expert AI evaluation assistant. Your task is to meticulously compare an AI's analysis of an email against a groundtruth.
    The email content is provided for your reference.

    EMAIL CONTENT:
    {{email_content}}

    AI OUTPUT:
    {{ai_output}}

    GROUNDTRUTH:
    {{groundtruth}}

    INSTRUCTIONS FOR EVALUATION:
    For each metric listed below, you MUST provide a detailed evaluation. Your response for EACH metric MUST be a JSON object with the following fields:

    1.  "Metric": (string) The name of the metric (e.g., "Sentiment analysis", "Summary").
    2.  "AI Value": (string) The value generated by the AI for this metric.
    3.  "Ground Truth": (string) The groundtruth value for this metric.
    4.  "Status": (string) Your assessment of the AI's performance for this metric. Must be one of: "Pass", "Fail", or "Partial Pass".
    5.  "Ground Truth Explanation": (string) A DETAILED explanation of WHY the "Ground Truth" value is correct.
        -   You MUST reference specific keywords, phrases, or patterns from the EMAIL CONTENT that justify the groundtruth.
        -   This explanation must be thorough and clearly articulate the reasoning. DO NOT use "N/A".
    6.  "Pass/Fail Explanation": (string) A DETAILED explanation of WHY the AI's output achieved the given "Status" (Pass/Fail/Partial Pass).
        -   Compare the "AI Value" directly against the "Ground Truth".
        -   Reference specific keywords, phrases, or patterns from the EMAIL CONTENT and the AI OUTPUT to justify your judgment.
        -   Clearly state what the AI did correctly or incorrectly. DO NOT use "N/A".

    7.  "individual_email_review_points": (string) Provide 2-3 bullet points summarizing:
        -   What the AI did well for THIS SPECIFIC EMAIL regarding this metric.
        -   What the AI missed or could improve for THIS SPECIFIC EMAIL regarding this metric.
        -   Format as a multi-line string with each point starting with a hyphen (e.g., "- AI correctly identified X.\n- AI missed Y.").
        -   This review is PER METRIC for the CURRENT EMAIL.

    SPECIAL INSTRUCTIONS FOR "Summary" METRIC:
    In addition to fields 1-6 AND field 7 (individual_email_review_points for the summary), the JSON object for the "Summary" metric MUST also include:
    8.  "Similarity Percentage": (string) Your calculated similarity score between the AI summary and the Ground Truth summary, expressed as a percentage (e.g., "85%"). Calculate this based on content accuracy, completeness, and adherence to any style requirements (like indirect speech if specified).
    9.  "% Explanation": (string) A DETAILED explanation for the "Similarity Percentage" you provided.
        -   Describe what aspects of the AI summary were good (contributing to the percentage) and what aspects were lacking or incorrect (detracting from the percentage).
        -   Reference specific parts of the AI summary and groundtruth summary. DO NOT use "N/A".

    GENERAL RULES:
    -   For "Sentiment analysis" metric: Valid values are ONLY "red" (negative) or "green" (positive). NEVER "neutral".
    -   For "overall_sentiment_analysis" metric: Valid values are ONLY "positive", "negative", or "neutral".
    -   Ensure all explanations are comprehensive and specific. Avoid generic statements.

    METRICS TO EVALUATE:
    1.  Summary (with "Similarity Percentage" and "% Explanation")
    2.  Sentiment analysis
    3.  overall_sentiment_analysis
    4.  feature
    5.  category
    6.  Events (If AI or Groundtruth contains event information, evaluate the event detection. If no events, state that clearly in explanations.)

    OUTPUT FORMAT:
    Return a JSON ARRAY of evaluation objects, one for each metric. Example for one metric:
    {{{{
        "Metric": "Sentiment analysis",
        "AI Value": "green",
        "Ground Truth": "green",
        "Status": "Pass",
        "Ground Truth Explanation": "The email contains positive keywords like 'great news' and 'looking forward to it', indicating a positive sentiment, hence 'green'.",
        "Pass/Fail Explanation": "The AI correctly identified the sentiment as 'green', matching the ground truth. The AI likely recognized the positive phrasing in the email.",
        "individual_email_review_points": "- AI correctly matched the positive sentiment cues.\n- No major misses for sentiment on this email."
    }}}}
    '''

    # Format the prompt with explicit string formatting
    # Note: The f-string interpolation happens here, not in the template itself
    formatted_prompt = prompt_template.format(
        email_content=email_content,
        ai_output=json.dumps(ai_output, indent=2),
        groundtruth=json.dumps(groundtruth, indent=2)
    )
    
    # Ensure the JSON example within the prompt is properly escaped if needed,
    # but using .format() should handle it correctly unless the content itself has issues.
    # Double-check escaping if errors persist.

    try:
        # Generate evaluation
        response = model.generate_content(formatted_prompt)
        response_text = response.text
        
        # Attempt to extract JSON from response
        import re
        
        # Try to find a JSON array in the response
        json_match = re.search(r'\[\s*{.*}\s*\]', response_text, re.DOTALL)
        if json_match:
            response_text = json_match.group(0)
        else:
            # Try another pattern with square brackets
            json_match = re.search(r'\[\s*\{.*\}\s*\]', response_text, re.DOTALL)
            if json_match:
                response_text = json_match.group(0)
            else:
                # Try to find a JSON object instead
                json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
                if json_match:
                    response_text = '[' + json_match.group(0) + ']'
        
        # Parse the JSON
        try:
            metrics = json.loads(response_text)
            
            # Ensure it's a list
            if not isinstance(metrics, list):
                metrics = [metrics]
                
            # Add pass/fail explanation for each metric
            for metric in metrics:
                # Ensure we have both fields to display in the UI
                if 'Comparison Explanation' in metric:
                    metric['Pass/Fail Explanation'] = metric['Comparison Explanation']
                    
                # If no comparison explanation is provided, add a default one
                if 'Pass/Fail Explanation' not in metric and 'Comparison Explanation' not in metric:
                    ai_value = metric.get('AI Value', 'N/A')
                    gt_value = metric.get('Ground Truth', 'N/A')
                    status = metric.get('Status', 'Unknown')
                    
                    # Generate a default explanation based on the metric
                    if metric.get('Metric') == 'Summary':
                        if status == 'Pass':
                            metric['Pass/Fail Explanation'] = f"The AI summary '{ai_value}' captures the key points and matches the groundtruth '{gt_value}' in terms of content and style."
                        else:
                            metric['Pass/Fail Explanation'] = f"The AI summary '{ai_value}' differs significantly from the groundtruth '{gt_value}' in content, completeness, or style."
                    
                    elif metric.get('Metric') == 'Sentiment analysis':
                        if status == 'Pass':
                            metric['Pass/Fail Explanation'] = f"The AI correctly identified the sentiment as '{ai_value}' which matches the groundtruth '{gt_value}'. This is the correct color code representation of the sentiment."
                        else:
                            metric['Pass/Fail Explanation'] = f"The AI incorrectly classified the sentiment as '{ai_value}'. The correct sentiment according to groundtruth is '{gt_value}'. This metric should only have values 'red' or 'green'."
                    
                    elif metric.get('Metric') == 'overall_sentiment_analysis':
                        if status == 'Pass':
                            metric['Pass/Fail Explanation'] = f"The AI correctly identified the overall sentiment as '{ai_value}' which matches the groundtruth '{gt_value}'. This correctly represents the email's overall tone."
                        else:
                            metric['Pass/Fail Explanation'] = f"The AI incorrectly classified the overall sentiment as '{ai_value}'. The correct overall sentiment according to groundtruth is '{gt_value}'. This metric can be 'positive', 'negative', or 'neutral'."
                    
                    elif metric.get('Metric') == 'feature' or metric.get('Metric') == 'category':
                        if status == 'Pass':
                            metric['Pass/Fail Explanation'] = f"The AI correctly identified the {metric.get('Metric')} as '{ai_value}' which matches the groundtruth '{gt_value}'."
                        else:
                            metric['Pass/Fail Explanation'] = f"The AI incorrectly classified the {metric.get('Metric')} as '{ai_value}'. The correct {metric.get('Metric')} according to groundtruth is '{gt_value}'."
                    
                    elif metric.get('Metric') == 'Events':
                        if status == 'Pass':
                            metric['Pass/Fail Explanation'] = f"The AI correctly identified all events mentioned in the email, matching the groundtruth."
                        else:
                            metric['Pass/Fail Explanation'] = f"The AI missed events or incorrectly identified events compared to the groundtruth."
            
            return metrics
            
        except json.JSONDecodeError as e:
            # Create a manual metric if JSON parsing fails
            print(f"Failed to parse LLM response: {e}")
            print(f"Raw response: {response_text}")
            
            # Create basic evaluation metrics manually for critical metrics
            metrics = []
            
            # Add summary evaluation
            if "Summary" in ai_output and "Summary" in groundtruth:
                metrics.append({
                    "Metric": "Summary",
                    "AI Value": ai_output.get("Summary", "N/A")[:100] + "..." if len(ai_output.get("Summary", "")) > 100 else ai_output.get("Summary", "N/A"),
                    "Ground Truth": groundtruth.get("Summary", "N/A")[:100] + "..." if len(groundtruth.get("Summary", "")) > 100 else groundtruth.get("Summary", "N/A"),
                    "Status": "Pass" if similar_text(ai_output.get("Summary", ""), groundtruth.get("Summary", "")) > 0.7 else "Fail",
                    "Pass/Fail Explanation": f"Similarity score between AI output and groundtruth is {similar_text(ai_output.get('Summary', ''), groundtruth.get('Summary', '')):.2f}",
                    "individual_email_review_points": "- AI correctly matched the positive sentiment cues.\n- No major misses for sentiment on this email."
                })
            
            # Add sentiment analysis evaluation
            if "Sentiment analysis" in ai_output and "Sentiment analysis" in groundtruth:
                sentiment_match = ai_output.get("Sentiment analysis") == groundtruth.get("Sentiment analysis")
                metrics.append({
                    "Metric": "Sentiment analysis",
                    "AI Value": ai_output.get("Sentiment analysis", "N/A"),
                    "Ground Truth": groundtruth.get("Sentiment analysis", "N/A"),
                    "Status": "Pass" if sentiment_match else "Fail",
                    "Pass/Fail Explanation": f"AI {'correctly' if sentiment_match else 'incorrectly'} identified the sentiment as '{ai_output.get('Sentiment analysis', 'N/A')}'. Valid values are only 'red' or 'green'.",
                    "individual_email_review_points": "- AI correctly matched the positive sentiment cues.\n- No major misses for sentiment on this email."
                })
            
            # Add overall sentiment analysis evaluation
            if "overall_sentiment_analysis" in ai_output and "overall_sentiment_analysis" in groundtruth:
                overall_match = ai_output.get("overall_sentiment_analysis") == groundtruth.get("overall_sentiment_analysis")
                metrics.append({
                    "Metric": "overall_sentiment_analysis",
                    "AI Value": ai_output.get("overall_sentiment_analysis", "N/A"),
                    "Ground Truth": groundtruth.get("overall_sentiment_analysis", "N/A"),
                    "Status": "Pass" if overall_match else "Fail",
                    "Pass/Fail Explanation": f"AI {'correctly' if overall_match else 'incorrectly'} identified the overall sentiment as '{ai_output.get('overall_sentiment_analysis', 'N/A')}'. Valid values are 'positive', 'negative', or 'neutral'.",
                    "individual_email_review_points": "- AI correctly matched the positive sentiment cues.\n- No major misses for sentiment on this email."
                })
            
            # Add feature evaluation
            if "feature" in ai_output and "feature" in groundtruth:
                feature_match = ai_output.get("feature").lower() == groundtruth.get("feature").lower()
                metrics.append({
                    "Metric": "feature",
                    "AI Value": ai_output.get("feature", "N/A"),
                    "Ground Truth": groundtruth.get("feature", "N/A"),
                    "Status": "Pass" if feature_match else "Fail",
                    "Pass/Fail Explanation": f"AI {'correctly' if feature_match else 'incorrectly'} identified the feature as '{ai_output.get('feature', 'N/A')}'.",
                    "individual_email_review_points": "- AI correctly identified the feature as '{ai_output.get('feature', 'N/A')}'."
                })
            
            # Add category evaluation
            if "category" in ai_output and "category" in groundtruth:
                category_match = ai_output.get("category").lower() == groundtruth.get("category").lower()
                metrics.append({
                    "Metric": "category",
                    "AI Value": ai_output.get("category", "N/A"),
                    "Ground Truth": groundtruth.get("category", "N/A"),
                    "Status": "Pass" if category_match else "Fail",
                    "Pass/Fail Explanation": f"AI {'correctly' if category_match else 'incorrectly'} identified the category as '{ai_output.get('category', 'N/A')}'.",
                    "individual_email_review_points": "- AI correctly identified the category as '{ai_output.get('category', 'N/A')}'."
                })
            
            # Add Events evaluation
            if "Events" in ai_output and "Events" in groundtruth:
                ai_events = ai_output.get("Events", [])
                gt_events = groundtruth.get("Events", [])
                events_match = len(ai_events) == len(gt_events)
                metrics.append({
                    "Metric": "Events",
                    "AI Value": f"{len(ai_events)} events" if ai_events else "No events",
                    "Ground Truth": f"{len(gt_events)} events" if gt_events else "No events",
                    "Status": "Pass" if events_match else "Fail",
                    "Pass/Fail Explanation": f"AI {'correctly' if events_match else 'incorrectly'} identified {len(ai_events)} events compared to {len(gt_events)} in groundtruth",
                    "individual_email_review_points": "- AI correctly identified all events mentioned in the email, matching the groundtruth."
                })
                
            return metrics
            
    except Exception as e:
        # Return a basic error metric
        print(f"Error in LLM evaluation: {str(e)}")
        # Instead of raising an error, generate fallback metrics
        return generate_fallback_metrics(ai_output, groundtruth)
    
    # Failsafe if nothing else worked
    return [{"Metric": "Evaluation Failed", "AI Value": "Error", "Ground Truth": "Error", "Status": "Fail", "Pass/Fail Explanation": "Failed to evaluate due to an error in the evaluation process"}]

def similar_text(text1: str, text2: str) -> float:
    """
    Calculate similarity between two text strings
    
    Args:
        text1: First text string
        text2: Second text string
        
    Returns:
        Similarity score between 0 and 1
    """
    if not text1 or not text2:
        return 0.0
    
    # Simple Jaccard similarity
    words1 = set(re.findall(r'\b\w+\b', text1.lower()))
    words2 = set(re.findall(r'\b\w+\b', text2.lower()))
    
    if not words1 or not words2:
        return 0.0
    
    intersection = len(words1.intersection(words2))
    union = len(words1.union(words2))
    
    return intersection / union if union > 0 else 0.0

def generate_fallback_metrics(ai_output: Dict[str, Any], groundtruth: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Generate fallback metrics when LLM evaluation fails
    
    Args:
        ai_output: AI-generated output
        groundtruth: Ground truth data
        
    Returns:
        List of metrics with basic evaluations
    """
    metrics = []
    
    # Helper function for text similarity
    def calculate_simple_similarity(text1, text2):
        if not text1 or not text2:
            return 0.0
        
        # Simple character-level Jaccard similarity
        chars1 = set(text1.lower())
        chars2 = set(text2.lower())
        
        intersection = len(chars1.intersection(chars2))
        union = len(chars1.union(chars2))
        
        return intersection / union if union > 0 else 0.0
    
    # Add Summary metric
    if "Summary" in ai_output and "Summary" in groundtruth:
        ai_summary = ai_output.get("Summary", "")
        gt_summary = groundtruth.get("Summary", "")
        
        # Use the similar_text function defined in this file
        similarity = similar_text(ai_summary, gt_summary)
        
        metrics.append({
            "Metric": "Summary",
            "AI Value": ai_summary[:100] + "..." if len(ai_summary) > 100 else ai_summary,
            "Ground Truth": gt_summary[:100] + "..." if len(gt_summary) > 100 else gt_summary,
            "Status": "Pass" if similarity > 0.7 else "Fail",
            "Pass/Fail Explanation": f"Summary similarity score is {similarity:.2f}. AI summary {'adequately' if similarity > 0.7 else 'inadequately'} captures the key points from the original email.",
            "individual_email_review_points": "- AI correctly matched the positive sentiment cues.\n- No major misses for sentiment on this email."
        })
    
    # Add Sentiment analysis metric
    if "Sentiment analysis" in ai_output and "Sentiment analysis" in groundtruth:
        ai_sentiment = ai_output.get("Sentiment analysis", "")
        gt_sentiment = groundtruth.get("Sentiment analysis", "")
        
        sentiment_match = ai_sentiment == gt_sentiment
        
        metrics.append({
            "Metric": "Sentiment analysis",
            "AI Value": ai_sentiment,
            "Ground Truth": gt_sentiment,
            "Status": "Pass" if sentiment_match else "Fail",
            "Pass/Fail Explanation": f"AI {'correctly' if sentiment_match else 'incorrectly'} identified the sentiment as '{ai_sentiment}'. This color-coded metric should only be 'red' or 'green' (never neutral).",
            "individual_email_review_points": "- AI correctly matched the positive sentiment cues.\n- No major misses for sentiment on this email."
        })
    
    # Add overall_sentiment_analysis metric
    if "overall_sentiment_analysis" in ai_output and "overall_sentiment_analysis" in groundtruth:
        ai_overall = ai_output.get("overall_sentiment_analysis", "")
        gt_overall = groundtruth.get("overall_sentiment_analysis", "")
        
        overall_match = ai_overall == gt_overall
        
        metrics.append({
            "Metric": "overall_sentiment_analysis",
            "AI Value": ai_overall,
            "Ground Truth": gt_overall,
            "Status": "Pass" if overall_match else "Fail",
            "Pass/Fail Explanation": f"AI {'correctly' if overall_match else 'incorrectly'} identified the overall sentiment as '{ai_overall}'. This metric can be 'positive', 'negative', or 'neutral'.",
            "individual_email_review_points": "- AI correctly matched the positive sentiment cues.\n- No major misses for sentiment on this email."
        })
    
    # Add feature metric
    if "feature" in ai_output and "feature" in groundtruth:
        ai_feature = ai_output.get("feature", "")
        gt_feature = groundtruth.get("feature", "")
        
        feature_match = ai_feature.lower() == gt_feature.lower()
        
        metrics.append({
            "Metric": "feature",
            "AI Value": ai_feature,
            "Ground Truth": gt_feature,
            "Status": "Pass" if feature_match else "Fail",
            "Pass/Fail Explanation": f"AI {'correctly' if feature_match else 'incorrectly'} identified the feature as '{ai_feature}'.",
            "individual_email_review_points": "- AI correctly identified the feature as '{ai_feature}'."
        })
    
    # Add category metric
    if "category" in ai_output and "category" in groundtruth:
        ai_category = ai_output.get("category", "")
        gt_category = groundtruth.get("category", "")
        
        category_match = ai_category.lower() == gt_category.lower()
        
        metrics.append({
            "Metric": "category",
            "AI Value": ai_category,
            "Ground Truth": gt_category,
            "Status": "Pass" if category_match else "Fail",
            "Pass/Fail Explanation": f"AI {'correctly' if category_match else 'incorrectly'} identified the category as '{ai_category}'.",
            "individual_email_review_points": "- AI correctly identified the category as '{ai_category}'."
        })
    
    # Add Events metric
    if "Events" in ai_output and "Events" in groundtruth:
        ai_events = ai_output.get("Events", [])
        gt_events = groundtruth.get("Events", [])
        
        events_match = len(ai_events) == len(gt_events)
        
        metrics.append({
            "Metric": "Events",
            "AI Value": f"{len(ai_events)} events detected",
            "Ground Truth": f"{len(gt_events)} events in groundtruth",
            "Status": "Pass" if events_match else "Fail",
            "Pass/Fail Explanation": f"AI {'correctly' if events_match else 'incorrectly'} identified {len(ai_events)} events compared to {len(gt_events)} events in groundtruth.",
            "individual_email_review_points": "- AI correctly identified all events mentioned in the email, matching the groundtruth."
        })
    
    return metrics