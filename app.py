import io
import json
import os
import tempfile
import time # For potential waits or logging
from typing import List, Dict, Any, Optional, BinaryIO

import pandas as pd
import streamlit as st
from fillpdf import fillpdfs
# Import OpenAI library
from openai import APIError, AuthenticationError, RateLimitError, OpenAI

# --- Configuration ---
OPENAI_MODEL: str = "gpt-3.5-turbo"
# Alternative models: "gpt-4-turbo-preview", "gpt-4" (more expensive)
# Models supporting guaranteed JSON mode: gpt-4-turbo-preview, gpt-3.5-turbo-1106 and later
PDF_FIELD_CACHE_TTL: int = 600 # Cache PDF fields for 10 minutes
EXCEL_CACHE_TTL: int = 600   # Cache Excel data for 10 minutes


# --- Core Helper Functions ---

@st.cache_data(ttl=PDF_FIELD_CACHE_TTL)
def get_pdf_fields(pdf_path: str) -> Optional[List[str]]:
    """Extracts fillable field names from a PDF using fillpdf."""
    st.write(f"Attempting to extract fields from: {pdf_path}") # Debug info
    try:
        # Try standard method first
        fields_dict = fillpdfs.get_form_fields(pdf_path, use_utf8=False) # Explicitly False first

        # Fallback if standard method returns empty (sometimes happens)
        if not fields_dict:
            st.warning(
                "`get_form_fields` returned empty. "
                "Trying alternative methods (may be slower)..."
            )
            try:
                # Try forcing utf8 encoding
                fields_dict = fillpdfs.get_form_fields(pdf_path, use_utf8=True)
            except Exception as utf8_e:
                st.warning(f"UTF-8 fallback failed: {utf8_e}")
                fields_dict = None # If fallback also fails

            if not fields_dict:
                st.error("Failed to extract fields even with fallback methods.")
                return None

        field_keys = list(fields_dict.keys())
        st.write(f"Extracted {len(field_keys)} fields.") # Debug info
        return field_keys

    except FileNotFoundError:
        st.error(f"Error: PDF file not found at path: {pdf_path}")
        return None
    except Exception as e:
        st.error(f"Error reading PDF fields: {e}")
        st.warning("Ensure the uploaded PDF is fillable and not corrupted.")
        # Check if pdftk seems accessible for diagnostics
        try:
            fillpdfs.check_output(['pdftk', '--version'])
            st.info("pdftk check command succeeded, error might be PDF-specific.")
        except Exception as pdftk_e:
            st.error(
                f"pdftk check failed: {pdftk_e}. Please ensure 'pdftk' "
                "is installed and in your system PATH."
            )
        return None


@st.cache_data(ttl=EXCEL_CACHE_TTL)
def get_excel_headers(excel_file_obj: BinaryIO) -> Optional[List[str]]:
    """Extracts headers from the first sheet of an Excel file."""
    try:
        # Read from the uploaded file object's buffer in memory
        excel_buffer = io.BytesIO(excel_file_obj.getvalue())
        # Reset buffer position just in case it was read before
        excel_buffer.seek(0)
        df = pd.read_excel(excel_buffer, sheet_name=0, nrows=0) # Read only headers
        return df.columns.tolist()
    except Exception as e:
        st.error(f"Error reading Excel headers: {e}")
        return None


@st.cache_data(ttl=EXCEL_CACHE_TTL)
def get_excel_first_row_data(excel_file_obj: BinaryIO) -> Optional[Dict[str, str]]:
    """Reads the first data row from the first sheet of an Excel file."""
    try:
        # Read from the uploaded file object's buffer in memory
        excel_buffer = io.BytesIO(excel_file_obj.getvalue())
        # Reset buffer position
        excel_buffer.seek(0)
        # Read header + first data row, keep header to align dict keys correctly
        df = pd.read_excel(excel_buffer, sheet_name=0, nrows=1, header=0)

        if df.empty:
            st.warning("Excel file appears to have headers but no data rows.")
            return None

        # Convert first row to dictionary {header: value}, handling potential NaT/NaN
        first_row_dict = df.iloc[0].to_dict()
        # Replace N/A with empty string, ensure all values are strings
        cleaned_dict = {
            k: "" if pd.isna(v) else str(v)
            for k, v in first_row_dict.items()
        }
        return cleaned_dict
    except Exception as e:
        st.error(f"Error reading Excel data row: {e}")
        return None


# --- LLM Mapping Function (OpenAI) ---

# @st.cache_data # Caching LLM calls can be complex due to variability; consider carefully
def map_fields_with_llm(
    pdf_fields: List[str],
    excel_headers: List[str]
) -> Optional[Dict[str, Optional[str]]]:
    """Uses OpenAI API to map PDF fields to Excel headers."""
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
        client = OpenAI(api_key=api_key)
    except KeyError:
        st.error("`OPENAI_API_KEY` not found in `secrets.toml`. Please add it.")
        st.code(
            "# .streamlit/secrets.toml\n"
            "OPENAI_API_KEY = \"sk-YOUR_API_KEY_HERE\"",
            language="toml"
        )
        return None
    except Exception as e:
        st.error(f"Error initializing OpenAI client: {e}")
        return None

    # Prepare the prompt messages for ChatCompletion
    system_prompt = (
        "You are an assistant that maps fields from a PDF form to columns "
        "in an Excel spreadsheet. Your output MUST be a single JSON object "
        "where keys are the PDF field names and values are the corresponding "
        "Excel column headers. If no suitable Excel column header is found for a "
        "PDF field, use a JSON null value for that key."
    )
    # Break long f-string for readability
    user_prompt = (
        "Analyze the following PDF field names and Excel column headers. "
        "Create a mapping from each PDF field name to the most appropriate "
        "Excel column header based on semantic meaning.\n\n"
        "PDF Field Names:\n"
        f"{json.dumps(pdf_fields, indent=2)}\n\n"
        "Excel Column Headers:\n"
        f"{json.dumps(excel_headers, indent=2)}\n\n"
        "Provide the mapping as a single JSON object. For PDF fields with "
        "no clear match in the Excel headers, map them to `null`."
    )

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]

    raw_json_response: Optional[str] = None
    try:
        st.write(f"### Attempting AI Field Mapping (using OpenAI {OPENAI_MODEL})...")
        with st.spinner("🤖 Asking AI to map fields... (This might take a moment)"):
            response = client.chat.completions.create(
                model=OPENAI_MODEL,
                messages=messages,
                response_format={"type": "json_object"}, # Request JSON mode
                temperature=0.1, # Lower temperature for more deterministic mapping
                max_tokens=2048, # Adjust if needed (check model limits)
            )
            raw_json_response = response.choices[0].message.content

        if not raw_json_response:
             st.error("❌ AI returned an empty response.")
             return None

        mapping: Dict[str, Optional[str]] = json.loads(raw_json_response)

        # --- Validation of the received mapping ---
        validated_mapping: Dict[str, Optional[str]] = {}
        missing_keys: List[str] = []
        invalid_value_mappings: Dict[str, Any] = {}
        valid_excel_headers_set = set(excel_headers)

        for field in pdf_fields:
            if field not in mapping:
                missing_keys.append(field)
                validated_mapping[field] = None # Treat missing key as unmapped
            else:
                mapped_value = mapping[field]
                # Check if the mapped value is valid (either None/null or an existing Excel header)
                if mapped_value is not None and mapped_value not in valid_excel_headers_set:
                    invalid_value_mappings[field] = mapped_value
                    validated_mapping[field] = None # Treat invalid mapping as unmapped
                else:
                    # Accept valid mapping (including null/None)
                    validated_mapping[field] = mapped_value

        st.success("AI mapping attempt complete. Validation finished.")

        # --- Display Mapping Results ---
        st.write("#### AI Mapping Results (Validated):")
        results_data = []
        unmapped_count = 0
        for pdf_field, mapped_excel_col in validated_mapping.items():
            status = "Mapped" if mapped_excel_col else "Not Mapped"
            if mapped_excel_col is None:
                unmapped_count += 1
            notes = []
            if pdf_field in missing_keys:
                notes.append("LLM failed to include this field.")
            if pdf_field in invalid_value_mappings:
                notes.append(
                    f"LLM suggested invalid header '{invalid_value_mappings[pdf_field]}'."
                )

            results_data.append({
                "PDF Field": pdf_field,
                "Mapped Excel Column": mapped_excel_col if mapped_excel_col else "*Unmapped*",
                "Status": status,
                "Notes": " ".join(notes) if notes else "-",
            })
        st.dataframe(results_data, use_container_width=True)

        # --- Report Unmapped PDF Fields ---
        if unmapped_count > 0:
            st.warning(
                f"{unmapped_count} PDF field(s) could not be mapped by the "
                "AI or required correction."
            )
            st.info("Unmapped fields will remain blank in the output PDF.")

        # --- Report Unused Excel Headers ---
        all_excel_headers_set = set(excel_headers)
        used_excel_headers_set = {
            h for h in validated_mapping.values() if h is not None
        }
        unused_excel_headers = sorted(
            list(all_excel_headers_set - used_excel_headers_set)
        )

        if unused_excel_headers:
            with st.expander(f"ℹ️ {len(unused_excel_headers)} Excel columns were not used for mapping"):
                st.write(
                    "These Excel columns were available but did not map to "
                    "any PDF field:"
                )
                st.json(unused_excel_headers)

        return validated_mapping # Return the validated mapping

    # --- Error Handling for LLM Call ---
    except json.JSONDecodeError:
        st.error("❌ AI response was not valid JSON. Cannot proceed with mapping.")
        st.text("Raw response received:")
        st.code(raw_json_response or "Could not capture raw response.", language="text")
        return None
    except AuthenticationError:
        st.error("❌ OpenAI Authentication Error: Check your API Key in secrets.toml.")
        return None
    except RateLimitError:
        st.error("❌ OpenAI Rate Limit Error: Quota exceeded or rate limit hit. Check OpenAI account.")
        return None
    except APIError as e:
        st.error(f"❌ OpenAI API Error: {e}")
        return None
    except Exception as e:
        st.error(f"❌ An unexpected error occurred during AI mapping: {e}")
        st.error(f"Error Type: {type(e).__name__}")
        return None


# --- PDF Filling Function ---

def fill_pdf_form(
    pdf_template_path: str,
    output_path: str,
    data_dict: Dict[str, str]
) -> bool:
    """Fills the PDF form using fillpdf and saves it."""
    try:
        # fillpdf generally expects strings. Ensure data_dict values are strings.
        # (Note: The calling code already does this conversion, but defense-in-depth)
        string_data_dict = {k: str(v) for k, v in data_dict.items()}

        # Use flatten=False to keep fields editable after filling if needed
        fillpdfs.write_fillable_pdf(
            pdf_template_path,
            output_path,
            string_data_dict,
            flatten=False
        )
        return True
    except Exception as e:
        st.error(f"Error filling PDF with pdftk: {e}")
        st.warning("This often indicates an issue with 'pdftk' execution.")
        # Check if pdftk seems accessible
        try:
            fillpdfs.check_output(['pdftk', '--version'])
            st.info(
                "pdftk check command succeeded, the error might be specific "
                "to the PDF or data."
            )
        except Exception as pdftk_e:
            st.error(
                f"pdftk check command failed: {pdftk_e}. Please ensure "
                "'pdftk' is installed and in your system PATH."
            )
        return False


# --- Streamlit App UI ---

def main():
    """Runs the Streamlit application."""
    st.set_page_config(layout="wide")
    st.title("📄➡️📊 AI PDF Filler from Excel (OpenAI-Powered)")
    st.markdown(f"""
    Upload a fillable PDF template and an Excel file (.xlsx).
    An AI ({OPENAI_MODEL}) will attempt to map Excel columns to PDF fields based on their names.
    The PDF will then be filled using data from the **first data row** of the Excel file according to the mapping.
    """)
    st.info("ℹ️ Requires **pdftk** to be installed and accessible system-wide.")
    st.info("ℹ️ Requires an **OpenAI API Key** stored in `.streamlit/secrets.toml`.")
    st.sidebar.warning(
        "🧪 AI mapping is experimental. Please **review the proposed mapping carefully** "
        "before generating the PDF."
    )

    # --- API Key Check ---
    api_key_present = "OPENAI_API_KEY" in st.secrets

    if not api_key_present:
        st.error(
            "OpenAI API Key (`OPENAI_API_KEY`) not found in `.streamlit/secrets.toml`. "
            "AI mapping is disabled."
        )
        st.code("""
# Create .streamlit/secrets.toml file with:
OPENAI_API_KEY = "sk-YOUR_OPENAI_API_KEY_HERE"
""", language="toml")

    # --- File Uploaders ---
    col1, col2 = st.columns(2)
    with col1:
        st.header("1. Upload PDF Template")
        uploaded_pdf: Optional[BinaryIO] = st.file_uploader(
            "Choose a fillable PDF file", type="pdf", key="pdf_uploader"
        )
    with col2:
        st.header("2. Upload Excel Data")
        uploaded_excel: Optional[BinaryIO] = st.file_uploader(
            "Choose an Excel file (.xlsx)", type="xlsx", key="excel_uploader"
        )

    # --- Main Processing Logic ---
    if uploaded_pdf and uploaded_excel:
        st.header("3. Process and Map")

        pdf_temp_path: Optional[str] = None
        output_pdf_path: Optional[str] = None
        mapping: Optional[Dict[str, Optional[str]]] = None
        pdf_fields: Optional[List[str]] = None
        excel_headers: Optional[List[str]] = None
        excel_data_first_row: Optional[Dict[str, str]] = None
        pdf_processed = False
        excel_processed = False

        try:
            # Create a temporary file for the PDF to get a stable path
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf_file:
                tmp_pdf_file.write(uploaded_pdf.getvalue())
                pdf_temp_path = tmp_pdf_file.name

            # --- Read Files ---
            # Pass the temp path for PDF, file object for Excel
            pdf_fields = get_pdf_fields(pdf_temp_path)
            excel_headers = get_excel_headers(uploaded_excel)
            excel_data_first_row = get_excel_first_row_data(uploaded_excel)

            # --- Check if file reading was successful ---
            pdf_processed = pdf_fields is not None
            # Check headers exist AND data row was read (even if empty dict initially)
            excel_processed = excel_headers is not None and excel_data_first_row is not None

            if pdf_processed and excel_processed:
                st.success("Successfully read PDF fields and Excel headers/data.")
                st.write(f"- Found {len(pdf_fields)} fillable fields in PDF.")
                st.write(f"- Found {len(excel_headers)} columns in Excel (first sheet).")

                # --- Perform AI Mapping ---
                if not api_key_present:
                    st.warning("Cannot perform AI mapping without OpenAI API Key.")
                else:
                    mapping = map_fields_with_llm(pdf_fields, excel_headers)

            # --- Provide feedback if reading failed ---
            elif not pdf_processed:
                st.error(
                    "Could not extract fields from the PDF. Is it a fillable form? "
                    "Check pdftk installation and PDF file validity."
                )
            elif not excel_headers:
                st.error("Could not read headers from the Excel file.")
            # Handles case where headers were read, but data read returned None
            elif excel_data_first_row is None and excel_headers:
                 st.error("Excel file has headers but no data could be read from the first row.")
            # General failure if excel_data_first_row is None (covers header failure too, but less specific)
            elif not excel_processed:
                 st.error("Could not read necessary data from the Excel file.")


            # --- Generate Button (only if mapping succeeded) ---
            if mapping:
                st.header("4. Generate Filled PDF")
                if st.button("Generate Filled PDF", key="generate_button", type="primary"):

                    # Prepare data dictionary for filling using the validated mapping
                    data_to_fill: Dict[str, str] = {}
                    skipped_fields_log: List[str] = []

                    # Ensure excel_data_first_row is not None before proceeding
                    if excel_data_first_row is None:
                         st.error("Cannot generate PDF: Excel data row is missing.")
                    else:
                        for pdf_field, excel_header in mapping.items():
                            if excel_header is None: # Skip unmapped fields
                                skipped_fields_log.append(f"'{pdf_field}' (Unmapped by AI)")
                                continue

                            # Get data, using mapped header as key for the excel data dict
                            if excel_header in excel_data_first_row:
                                value = excel_data_first_row[excel_header]
                                # Value should already be string or empty string from get_excel_first_row_data
                                data_to_fill[pdf_field] = value
                            else:
                                # Should be rare if excel_data_first_row was read correctly and headers match
                                st.warning(
                                    f"Mapped Excel header '{excel_header}' for PDF field "
                                    f"'{pdf_field}' not found in the first data row dictionary. Skipping."
                                )
                                skipped_fields_log.append(
                                    f"'{pdf_field}' (Header '{excel_header}' missing in data)"
                                )

                        if skipped_fields_log:
                            st.info(f"Skipped filling fields: {', '.join(skipped_fields_log)}")

                        if not data_to_fill:
                            st.error("No data could be prepared for filling based on the AI mapping and available Excel data.")
                        else:
                            st.write("##### Data to be Filled:")
                            # Show the final data being sent to fillpdf
                            st.json(data_to_fill)

                            try:
                                with st.spinner("⏳ Generating filled PDF..."):
                                    # Create a temporary file for the output PDF
                                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_output_pdf:
                                        output_pdf_path = tmp_output_pdf.name

                                    # Fill the form using the input temp path and output temp path
                                    success = fill_pdf_form(
                                        pdf_temp_path,
                                        output_pdf_path,
                                        data_to_fill
                                    )

                                if success and output_pdf_path:
                                    st.success("✅ PDF Filled Successfully!")
                                    # Read the generated PDF into memory for download
                                    try:
                                        with open(output_pdf_path, "rb") as f:
                                            pdf_bytes = f.read()

                                        download_filename = f"filled_{uploaded_pdf.name}"
                                        st.download_button(
                                            label="⬇️ Download Filled PDF",
                                            data=pdf_bytes,
                                            file_name=download_filename,
                                            mime="application/pdf",
                                        )
                                    except FileNotFoundError:
                                         st.error("Failed to read the generated PDF file for download.")
                                    except Exception as read_err:
                                         st.error(f"An error occurred preparing the download: {read_err}")
                                # No else needed for `if success`, fill_pdf_form shows errors

                            finally:
                                # Clean up the output temporary file if it was created
                                if output_pdf_path and os.path.exists(output_pdf_path):
                                    try:
                                        os.remove(output_pdf_path)
                                    except PermissionError:
                                        st.info("Could not immediately delete temporary output PDF (might be locked).")
                                    except Exception as e:
                                        st.warning(f"Could not delete temporary output file '{output_pdf_path}': {e}")

            # --- Show message if mapping failed but files were okay ---
            elif pdf_processed and excel_processed and api_key_present:
                st.error("PDF Generation blocked: AI mapping failed or produced no valid results.")
            elif pdf_processed and excel_processed and not api_key_present:
                 # No button shown, but indicate files were processed
                 st.info("Files processed. Add API key to enable AI mapping and generation.")


        finally:
            # --- Clean up the input temporary PDF file ---
            if pdf_temp_path and os.path.exists(pdf_temp_path):
                try:
                    os.remove(pdf_temp_path)
                except PermissionError:
                    # On Windows, files might be temporarily locked
                    st.info(
                        "Could not immediately delete temporary input PDF file (might be locked)."
                    )
                except Exception as e:
                    st.warning(f"Could not delete temporary input file '{pdf_temp_path}': {e}")

    # --- UI Prompts if files are missing ---
    elif uploaded_pdf and not uploaded_excel:
        st.warning("Please upload the Excel data file.")
    elif not uploaded_pdf and uploaded_excel:
        st.warning("Please upload the PDF template file.")
    else:
        # Initial state, no files uploaded yet
        st.info("Upload a fillable PDF and an Excel file to start.")


if __name__ == "__main__":
    main()
