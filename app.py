import streamlit as st
import pandas as pd
import pdfrw # For reading/writing PDF form fields
import json
import io
import os
from mistralai import Mistral
import tempfile # To handle uploaded files safely


# --- Configuration ---
MODEL_NAME = "pixtral-12b-latest"
DEFAULT_PROMPT_TEMPLATE = """
You are an AI assistant tasked with matching field names from an Excel sheet to field names in a fillable PDF form.
The goal is to determine which Excel field likely corresponds to which PDF field, even if the names aren't exact matches (e.g., "Company Name" vs "Business Name").

Here are the available PDF field names:
{pdf_fields_list}

Here are the available Excel field names (these have corresponding values):
{excel_fields_list}

Please provide a JSON object mapping *PDF field names* to the *most relevant Excel field name*.
- The keys of the JSON object should be the PDF field names.
- The values should be the corresponding Excel field names.
- If you cannot find a reasonable match for a PDF field, DO NOT include it in the JSON output.
- Ensure the output is ONLY the JSON object, nothing else.

Example Output Format:
{{
  "PDF Field Name 1": "Excel Field Name A",
  "PDF Field Name 2": "Excel Field Name B"
  # ... only include matched fields
}}

JSON mapping:
"""

# --- Helper Functions ---

def get_pdf_fields(pdf_bytes_io):
    """Reads field names from a fillable PDF using pdfrw."""
    fields = {}
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
        if pdf.Root and '/AcroForm' in pdf.Root and '/Fields' in pdf.Root.AcroForm:
            for field in pdf.Root.AcroForm.Fields:
                # Field names are usually under /T key, remove surrounding parentheses if present
                field_name = field.get('/T')
                if field_name:
                    field_name = field_name.strip('()')
                    # Store the raw field object along with the name for later filling
                    fields[field_name] = field
                    st.write(field_name)
        return fields # Return dict {field_name: field_object}
    except Exception as e:
        st.error(f"Error reading PDF fields: {e}")
        # Try to read annotations directly if AcroForm fails (less reliable)
        try:
            pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
            alt_fields = {}
            for page in pdf.pages:
                 if page.Annots:
                     for annot in page.Annots:
                         if annot.Subtype == '/Widget' and annot.T:
                            field_name = annot.T.strip('()')
                            alt_fields[field_name] = annot # Store annot object
            if alt_fields:
                 st.warning("Used alternative method to find fields. Results may vary.")
                 return alt_fields
            else:
                 st.error("Could not find any form fields using standard or alternative methods.")
                 return {} # Return empty if absolutely nothing found
        except Exception as e_alt:
            st.error(f"Further error during alternative PDF field reading: {e_alt}")
            return {} # Return empty on secondary error

def read_excel_data(excel_bytes_io):
    """Reads data from the first two columns of an Excel file."""
    try:
        df = pd.read_excel(excel_bytes_io, header=None, usecols=[0, 1], engine='openpyxl')
        # Convert empty/NaN keys to a placeholder or skip them
        df = df.dropna(subset=[0]) # Drop rows where the field name (col 0) is empty
        df[0] = df[0].astype(str) # Ensure field names are strings
        df[1] = df[1].fillna('').astype(str) # Ensure values are strings, fill NaN with empty string
        data_dict = dict(zip(df[0], df[1]))
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None

def match_fields_with_ai(pdf_field_names, excel_field_names):
    """Uses Mistral API to match PDF fields to Excel fields."""
    api_key = st.secrets["MISTRAL_API_KEY"]
    if not api_key:
        st.error("Mistral API Key is required.") # Updated error message
        return None
    if not pdf_field_names or not excel_field_names:
        st.warning("Cannot perform matching without both PDF and Excel fields.")
        return {} # Return empty mapping if no fields

    # Instantiate Mistral client
    client = Mistral(api_key=api_key)

    # Use the same prompt template
    prompt = DEFAULT_PROMPT_TEMPLATE.format(
        pdf_fields_list="\n".join([f"- {f}" for f in pdf_field_names]),
        excel_fields_list="\n".join([f"- {f}" for f in excel_field_names])
    )

    # Format messages for Mistral API
    messages = [
         {
                "role": "system",
                "content": "You are an expert at comparing text input field names.",
            },
            {
                "role": "user",
                "content": f"{prompt}",
            },
        ]

    try:
        st.info(f"Asking {MODEL_NAME} to match fields...") # Updated info message
        # Call Mistral chat endpoint
        completion = client.chat.complete(
            model=MODEL_NAME,
            messages=messages,
            temperature=0.1, # Lower temperature for more deterministic mapping
            response_format={"type": "json_object"} # Request JSON output directly
        )
        response_content = completion.choices[0].message.content
        st.write("AI Response (Raw JSON):")
        st.code(response_content, language='json') # Show the raw response for debugging

        # Attempt to parse the JSON response
        mapping = json.loads(response_content)

        # --- Validation Step (remains the same) ---
        validated_mapping = {}
        valid_pdf_keys = set(pdf_field_names)
        valid_excel_values = set(excel_field_names)
        for pdf_key, excel_val in mapping.items():
            if pdf_key in valid_pdf_keys and excel_val in valid_excel_values:
                 validated_mapping[pdf_key] = excel_val
            else:
                 st.warning(f"AI proposed an invalid mapping - PDF:'{pdf_key}' -> Excel:'{excel_val}'. Skipping.")
        # ----------------------

        st.success(f"{MODEL_NAME} matching complete.") # Updated success message
        return validated_mapping # Return the validated mapping

    except json.JSONDecodeError as e:
        st.error(f"Error parsing AI response as JSON: {e}")
        st.error(f"Raw response was: {response_content}")
        return None
    except Exception as e:
        # Catch potential Mistral API specific errors if needed, otherwise generic Exception
        st.error(f"Error calling Mistral API: {e}")
        return None

def fill_pdf(pdf_bytes_io, field_mapping, excel_data, pdf_fields_objects):
    """Fills the PDF form fields based on the mapping and data."""
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())

        # Ensure NeedAppearances is set for viewers to render fields correctly
        if pdf.Root.AcroForm:
             pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))

        filled_count = 0
        for pdf_field_name, excel_field_name in field_mapping.items():
            if pdf_field_name in pdf_fields_objects and excel_field_name in excel_data:
                field_obj = pdf_fields_objects[pdf_field_name]
                value_to_fill = excel_data[excel_field_name]

                # Update the field value (/V) and potentially appearance (/AP)
                # Using PdfString ensures correct PDF encoding
                from pdfrw.objects import pdfstring
                field_obj.update(pdfrw.PdfDict(V=pdfstring.PdfString.encode(value_to_fill)))
                # Optionally clear the appearance stream (/AP) so the viewer regenerates it
                if '/AP' in field_obj:
                  del field_obj['/AP']
                filled_count += 1
            else:
                 st.warning(f"Skipping field '{pdf_field_name}': Mapped Excel field '{excel_field_name}' not found in data or PDF field object missing.")

        st.info(f"Attempted to fill {filled_count} fields based on mapping.")

        # Write the modified PDF to a BytesIO object
        output_pdf_stream = io.BytesIO()
        pdfrw.PdfWriter().write(output_pdf_stream, pdf)
        output_pdf_stream.seek(0) # Rewind the stream to the beginning
        return output_pdf_stream

    except Exception as e:
        st.error(f"Error filling PDF: {e}")
        return None

# --- Streamlit UI ---

st.set_page_config(layout="wide")
st.title("📄➡️📊 AI PDF Form Filler")
st.markdown("Upload a fillable PDF and an Excel file (.xlsx). The app uses AI to match Excel data (Col A: Field Name, Col B: Value) to PDF fields and fills the form.")

# Use columns for better layout
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Files")
    uploaded_pdf = st.file_uploader("Upload Fillable PDF", type="pdf")
    uploaded_excel = st.file_uploader("Upload Excel Data (.xlsx)", type="xlsx")

st.subheader("3. Process & Download")
fill_button = st.button("✨ Fill PDF using AI Matcher")

# --- Main Logic ---
if fill_button:
    if uploaded_pdf and uploaded_excel:
        # Process in memory using BytesIO
        pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())
        excel_bytes_io = io.BytesIO(uploaded_excel.getvalue())

        pdf_fields_objects = {}
        pdf_field_names = []
        excel_data = {}

        with st.spinner("Reading PDF fields..."):
            pdf_fields_objects = get_pdf_fields(pdf_bytes_io) # Returns {name: object}
            pdf_field_names = list(pdf_fields_objects.keys()) # Get just the names for matching
            if pdf_field_names:
                st.write("Detected PDF Fields:")
                st.dataframe(pdf_field_names, use_container_width=True)
            else:
                st.error("No fillable fields detected in the PDF.")
                st.stop() # Stop execution if no fields

        with st.spinner("Reading Excel data..."):
            excel_data = read_excel_data(excel_bytes_io) # Returns {name: value}
            if excel_data:
                st.write("Detected Excel Data (Field -> Value):")
                # Convert dict to DataFrame for better display
                excel_df_display = pd.DataFrame(list(excel_data.items()), columns=['Excel Field', 'Value'])
                st.dataframe(excel_df_display, use_container_width=True)
            else:
                st.error("Could not read data from the Excel file.")
                st.stop() # Stop execution if no data

        # Perform AI Matching
        field_mapping = None
        with st.spinner(f"Matching fields with {MODEL_NAME}..."):
            field_mapping = match_fields_with_ai(pdf_field_names, list(excel_data.keys()))

        if field_mapping:
            st.write("AI Field Mapping (PDF Field -> Excel Field):")
            # Convert mapping dict to DataFrame for display
            map_df_display = pd.DataFrame(list(field_mapping.items()), columns=['PDF Field', 'Matched Excel Field'])
            st.dataframe(map_df_display, use_container_width=True)

            # Fill the PDF
            with st.spinner("Filling PDF form..."):
                # Rewind the PDF stream before filling
                pdf_bytes_io.seek(0)
                filled_pdf_stream = fill_pdf(pdf_bytes_io, field_mapping, excel_data, pdf_fields_objects)

            if filled_pdf_stream:
                st.success("PDF Filled Successfully!")
                # Provide download button
                output_filename = f"filled_{uploaded_pdf.name}"
                st.download_button(
                    label="⬇️ Download Filled PDF",
                    data=filled_pdf_stream,
                    file_name=output_filename,
                    mime="application/pdf"
                )
        elif field_mapping == {}: # Case where AI found zero matches
             st.warning("The AI could not confidently match any PDF fields to the Excel data provided.")
        else:
            st.error("PDF filling could not proceed due to issues in the AI matching step.")

    else:
        # Error messages if files/key are missing
        if not uploaded_pdf:
            st.warning("Please upload a PDF file.")
        if not uploaded_excel:
            st.warning("Please upload an Excel file.")
        if not api_key:
            st.warning("Please enter your OpenAI API Key.")
