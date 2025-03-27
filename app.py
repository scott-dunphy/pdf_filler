import streamlit as st
import pandas as pd
import pdfrw  # For reading/writing PDF form fields
import json
import io
import os
from mistralai import Mistral
import tempfile  # To handle uploaded files safely

# --- Configuration ---
MODEL_NAME = "pixtral-12b-latest"
DEFAULT_PROMPT_TEMPLATE = """
You are an AI assistant tasked with matching field names from an Excel sheet to field names in a fillable PDF form.
The goal is to determine which Excel field likely corresponds to which PDF field, even if the names aren't exact matches (e.g., "Company Name" vs "Business Name").
Be careful not to duplicate values like names if one is First Name and the other Spouse's First Name.

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
                field_name = field.get('/T')
                if field_name:
                    field_name = field_name.strip('()')
                    fields[field_name] = field
                    st.write(f"Detected field: {field_name}")
        return fields  # Return dict {field_name: field_object}
    except Exception as e:
        st.error(f"Error reading PDF fields: {e}")
        try:
            pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
            alt_fields = {}
            for page in pdf.pages:
                if page.Annots:
                    for annot in page.Annots:
                        if annot.Subtype == '/Widget' and annot.T:
                            field_name = annot.T.strip('()')
                            alt_fields[field_name] = annot
            if alt_fields:
                st.warning("Used alternative method to find fields. Results may vary.")
                return alt_fields
            else:
                st.error("Could not find any form fields using standard or alternative methods.")
                return {}
        except Exception as e_alt:
            st.error(f"Further error during alternative PDF field reading: {e_alt}")
            return {}

def read_excel_data(excel_bytes_io):
    """Reads data from the first two columns of an Excel file."""
    try:
        df = pd.read_excel(excel_bytes_io, header=None, usecols=[0, 1], engine='openpyxl')
        df = df.dropna(subset=[0])
        df[0] = df[0].astype(str)
        df[1] = df[1].fillna('').astype(str)
        data_dict = dict(zip(df[0], df[1]))
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None

def match_fields_with_ai(pdf_field_names, excel_field_names):
    """Uses Mistral API to match PDF fields to Excel fields."""
    api_key = st.secrets["MISTRAL_API_KEY"]
    if not api_key:
        st.error("Mistral API Key is required.")
        return None
    if not pdf_field_names or not excel_field_names:
        st.warning("Cannot perform matching without both PDF and Excel fields.")
        return {}

    client = Mistral(api_key=api_key)
    prompt = DEFAULT_PROMPT_TEMPLATE.format(
        pdf_fields_list="\n".join([f"- {f}" for f in pdf_field_names]),
        excel_fields_list="\n".join([f"- {f}" for f in excel_field_names])
    )
    messages = [
        {"role": "system", "content": "You are an expert at comparing text input field names."},
        {"role": "user", "content": f"{prompt}"},
    ]
    try:
        st.info(f"Asking {MODEL_NAME} to match fields...")
        completion = client.chat.complete(
            model=MODEL_NAME,
            messages=messages,
            temperature=0.1,
            response_format={"type": "json_object"}
        )
        response_content = completion.choices[0].message.content
        st.write("AI Response (Raw JSON):")
        st.code(response_content, language='json')
        mapping = json.loads(response_content)
        validated_mapping = {}
        valid_pdf_keys = set(pdf_field_names)
        valid_excel_values = set(excel_field_names)
        for pdf_key, excel_val in mapping.items():
            if pdf_key in valid_pdf_keys and excel_val in valid_excel_values:
                validated_mapping[pdf_key] = excel_val
            else:
                st.warning(f"AI proposed an invalid mapping - PDF:'{pdf_key}' -> Excel:'{excel_val}'. Skipping.")
        st.success(f"{MODEL_NAME} matching complete.")
        return validated_mapping
    except json.JSONDecodeError as e:
        st.error(f"Error parsing AI response as JSON: {e}")
        st.error(f"Raw response was: {response_content}")
        return None
    except Exception as e:
        st.error(f"Error calling Mistral API: {e}")
        return None

def update_field(field_obj, encoded_value):
    """Update a field object (and its kids) with the encoded value and remove the appearance stream."""
    field_obj.update(pdfrw.PdfDict(V=encoded_value, DV=encoded_value))
    if '/AP' in field_obj:
        del field_obj['/AP']
    if '/Kids' in field_obj:
        for kid in field_obj.Kids:
            kid.update(pdfrw.PdfDict(V=encoded_value, DV=encoded_value))
            if '/AP' in kid:
                del kid['/AP']

def fill_pdf(pdf_bytes_io, field_mapping, excel_data, pdf_fields_objects):
    """Fills the PDF form fields based on the mapping and data, with extra debugging and a second pass for widget annotations."""
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
        # Set NeedAppearances so the viewer will rebuild appearances
        if pdf.Root.AcroForm:
            pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))

        filled_count = 0
        # First pass: update using our known field objects
        for pdf_field_name, excel_field_name in field_mapping.items():
            if pdf_field_name in pdf_fields_objects and excel_field_name in excel_data:
                field_obj = pdf_fields_objects[pdf_field_name]
                value_to_fill = excel_data[excel_field_name]
                # Debug the value coming from Excel
                st.write(f"Updating field '{pdf_field_name}' with value '{value_to_fill}'")
                encoded_value = pdfrw.objects.pdfstring.PdfString.encode(value_to_fill)
                update_field(field_obj, encoded_value)
                filled_count += 1
            else:
                st.warning(f"Skipping field '{pdf_field_name}': Mapped Excel field '{excel_field_name}' not found or PDF field missing.")

        # Second pass: update any widget annotations in each page that match our mapping
        for page in pdf.pages:
            if page.Annots:
                for annot in page.Annots:
                    if annot.Subtype == '/Widget' and annot.T:
                        annot_field_name = annot.T.strip('()')
                        if annot_field_name in field_mapping:
                            excel_field_name = field_mapping[annot_field_name]
                            if excel_field_name in excel_data:
                                value_to_fill = excel_data[excel_field_name]
                                st.write(f"(Widget pass) Updating field '{annot_field_name}' with value '{value_to_fill}'")
                                encoded_value = pdfrw.objects.pdfstring.PdfString.encode(value_to_fill)
                                annot.update(pdfrw.PdfDict(V=encoded_value, DV=encoded_value))
                                if '/AP' in annot:
                                    del annot['/AP']

        st.info(f"Attempted to fill {filled_count} fields based on mapping.")

        # Debug: re-read the field values from pdf object before writing
        st.write("Debug: Field values after updating (AcroForm):")
        if pdf.Root.AcroForm and '/Fields' in pdf.Root.AcroForm:
            for field in pdf.Root.AcroForm.Fields:
                field_name = field.get('/T')
                field_value = field.get('/V')
                st.write(f"Field: {field_name}, Value: {field_value}")
        else:
            st.write("No AcroForm Fields found in the updated PDF.")

        # Write the modified PDF to a BytesIO object
        output_pdf_stream = io.BytesIO()
        pdfrw.PdfWriter().write(output_pdf_stream, pdf)
        output_pdf_stream.seek(0)
        return output_pdf_stream
    except Exception as e:
        st.error(f"Error filling PDF: {e}")
        return None

# --- Streamlit Main Code ---

if __name__ == "__main__":
    pass  # The main code is in the code block above.
