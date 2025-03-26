import streamlit as st
import pandas as pd
from fillpdf import fillpdfs
import tempfile
import os
import io
import json
# Import OpenAI library
from openai import OpenAI, APIError, AuthenticationError, RateLimitError
import time # For potential waits or logging

# --- Configuration ---
OPENAI_MODEL = "gpt-3.5-turbo"
# Alternative models: "gpt-4-turbo-preview", "gpt-4" (more expensive)
# Models supporting guaranteed JSON mode: gpt-4-turbo-preview, gpt-3.5-turbo-1106 and later

# --- Core Helper Functions ---

@st.cache_data(ttl=600) # Cache for 10 mins to avoid re-reading same file quickly
def get_pdf_fields(pdf_path):
    """Extracts fillable field names from a PDF."""
    try:
        # Try standard method first
        fields_dict = fillpdfs.get_form_fields(pdf_path)

        # Fallback if standard method returns empty (sometimes happens)
        if not fields_dict:
             st.warning("`get_form_fields` returned empty. Trying alternative methods (may be slower)...")
             try:
                 # Try forcing utf8 encoding
                 fields_dict = fillpdfs.get_form_fields(pdf_path, use_utf8=True)
             except Exception:
                 fields_dict = None # If fallback also fails

             if not fields_dict:
                 st.error("Failed to extract fields even with fallback methods.")
                 return None

        return list(fields_dict.keys())

    except Exception as e:
        st.error(f"Error reading PDF fields: {e}")
        st.warning("Ensure the uploaded PDF is fillable and not corrupted.")
        # Check if pdftk seems accessible
        try:
            fillpdfs.check_output(['pdftk', '--version'])
        except Exception as pdftk_e:
            st.error(f"pdftk check failed: {pdftk_e}. Please ensure 'pdftk' is installed and in your system PATH.")
        return None

@st.cache_data(ttl=600)
def get_excel_headers(excel_file_obj):
    """Extracts headers from the first sheet of an Excel file."""
    try:
        # Read from the uploaded file object's buffer in memory
        excel_buffer = io.BytesIO(excel_file_obj.getvalue())
        df = pd.read_excel(excel_buffer, sheet_name=0, nrows=0) # Read only headers
        return df.columns.tolist()
    except Exception as e:
        st.error(f"Error reading Excel headers: {e}")
        return None

@st.cache_data(ttl=600)
def get_excel_first_row_data(excel_file_obj):
    """Reads the first data row from the first sheet of an Excel file."""
    try:
        # Read from the uploaded file object's buffer in memory
        excel_buffer = io.BytesIO(excel_file_obj.getvalue())
        # Read header + first data row, keep header to align dict keys correctly
        df = pd.read_excel(excel_buffer, sheet_name=0, nrows=1, header=0)
        if df.empty:
            st.warning("Excel file appears to have headers but no data rows.")
            return None
        # Convert first row to dictionary {header: value}, handling potential NaT/NaN
        first_row_dict = df.iloc[0].to_dict()
        cleaned_dict = {}
        for k, v in first_row_dict.items():
            cleaned_dict[k] = "" if pd.isna(v) else str(v) # Replace N/A with empty string
        return cleaned_dict
    except Exception as e:
        st.error(f"Error reading Excel data row: {e}")
        return None

# --- LLM Mapping Function (OpenAI) ---
# @st.cache_data # Caching LLM calls can be complex due to variability; omit for now
def map_fields_with_llm(pdf_fields, excel_headers):
    """Uses OpenAI API to map PDF fields to Excel headers."""
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
        client = OpenAI(api_key=api_key)
    except KeyError:
        st.error("`OPENAI_API_KEY` not found in `secrets.toml`. Please add it.")
        st.code("# .streamlit/secrets.toml\nOPENAI_API_KEY = \"sk-YOUR_API_KEY_HERE\"", language="toml")
        return None
    except Exception as e:
        st.error(f"Error initializing OpenAI client: {e}")
        return None

    # Prepare the prompt messages for ChatCompletion
    system_prompt = "You are an assistant that maps fields from a PDF form to columns in an Excel spreadsheet. Your output must be a single JSON object."
    user_prompt = f"""
Analyze the following PDF field names and Excel column headers. Create a mapping from each PDF field name to the most appropriate Excel column header based on semantic meaning.

PDF Field Names:
```json
{json.dumps(pdf_fields, indent=2)}
