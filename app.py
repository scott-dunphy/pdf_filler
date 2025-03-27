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
    """Reads field names and types from a fillable PDF using pdfrw."""
    fields_data = {}  # Will store {field_name: {'obj': field_object, 'type': field_type}}
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
        if pdf.Root and '/AcroForm' in pdf.Root and '/Fields' in pdf.Root.AcroForm:
            for field in pdf.Root.AcroForm.Fields:
                field_name = field.get('/T')
                if field_name:
                    field_name = field_name.strip('()')
                    field_type = field.get('/FT')  # Field type (e.g., /Tx for text, /Btn for button/checkbox)
                    on_value = None
                    
                    # For checkboxes, try to determine the "On" value
                    if field_type == pdfrw.PdfName('Btn'):
                        # Look for the ON value in appearance states
                        if '/AP' in field and '/N' in field['/AP']:
                            ap_n = field['/AP']['/N']
                            if isinstance(ap_n, pdfrw.PdfDict):
                                # The key that is not '/Off' is likely the 'On' state
                                for key in ap_n.keys():
                                    key_str = str(key)
                                    if key_str != '/Off':
                                        on_value = pdfrw.PdfName(key_str.strip('/'))
                        
                        # If we couldn't find it in appearances, default to 'Yes'
                        if not on_value:
                            on_value = pdfrw.PdfName('Yes')
                        
                        st.write(f"Detected checkbox field: {field_name} (On value: {on_value})")
                    else:
                        st.write(f"Detected field: {field_name} (Type: {field_type})")
                    
                    fields_data[field_name] = {
                        'obj': field,
                        'type': field_type,
                        'on_value': on_value
                    }
            
            return fields_data
        else:
            st.warning("No AcroForm found in PDF Root. Trying alternative annotation search.")
            
    except Exception as e:
        st.error(f"Error reading PDF AcroForm fields: {e}")
    
    # Alternative method: Try to find fields using annotations
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
        alt_fields = {}
        for page in pdf.pages:
            if page.Annots:
                for annot in page.Annots:
                    if annot.Subtype == pdfrw.PdfName('Widget') and annot.T:
                        field_name = annot.T.strip('()')
                        field_type = annot.get('/FT')
                        on_value = None
                        
                        # For checkboxes, try to determine the "On" value
                        if field_type == pdfrw.PdfName('Btn'):
                            # Look for the ON value in appearance states
                            if '/AP' in annot and '/N' in annot['/AP']:
                                ap_n = annot['/AP']['/N']
                                if isinstance(ap_n, pdfrw.PdfDict):
                                    # The key that is not '/Off' is likely the 'On' state
                                    for key in ap_n.keys():
                                        key_str = str(key)
                                        if key_str != '/Off':
                                            on_value = pdfrw.PdfName(key_str.strip('/'))
                            
                            # If we couldn't find it in appearances, default to 'Yes'
                            if not on_value:
                                on_value = pdfrw.PdfName('Yes')
                            
                            st.write(f"(Annot) Detected checkbox field: {field_name} (On value: {on_value})")
                        else:
                            st.write(f"(Annot) Detected field: {field_name} (Type: {field_type})")
                        
                        alt_fields[field_name] = {
                            'obj': annot,
                            'type': field_type,
                            'on_value': on_value
                        }
        
        if alt_fields:
            st.warning("Used alternative method (page annotations) to find fields. Results may vary.")
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
    api_key = st.secrets.get("MISTRAL_API_KEY")  # Use .get for safer access
    if not api_key:
        st.error("Mistral API Key is required. Please set it in Streamlit secrets.")
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

        # Basic JSON parsing
        try:
            mapping = json.loads(response_content)
            if not isinstance(mapping, dict):
                st.error(f"AI response was not a JSON object (dictionary). Got type: {type(mapping)}")
                return None
        except json.JSONDecodeError as e:
            st.error(f"Error parsing AI response as JSON: {e}")
            st.error(f"Raw response was: {response_content}")
            # Attempt to extract JSON manually if possible (simple cases)
            try:
                start = response_content.find('{')
                end = response_content.rfind('}') + 1
                if start != -1 and end != -1:
                    cleaned_response = response_content[start:end]
                    mapping = json.loads(cleaned_response)
                    st.warning("Manually extracted JSON from AI response.")
                else:
                    raise ValueError("Could not find JSON object delimiters.")
            except Exception as extract_err:
                st.error(f"Could not manually extract JSON: {extract_err}")
                return None

        # Validation
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
    except Exception as e:
        st.error(f"Error calling Mistral API or processing response: {e}")
        import traceback
        st.error(traceback.format_exc())  # More detailed error for debugging
        return None

def fill_pdf(pdf_bytes_io, field_mapping, excel_data, pdf_fields_data):
    """Fills the PDF form fields based on the mapping and data."""
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
        # Set NeedAppearances so the viewer will rebuild appearances
        if pdf.Root and pdf.Root.AcroForm:
            pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        else:
            st.warning("PDF does not contain an AcroForm dictionary. Field appearances might not update correctly.")

        filled_count = 0
        processed_pdf_fields = set()  # Keep track of fields updated via AcroForm/Annotations

        # Iterate through the AI mapping
        for pdf_field_name, excel_field_name in field_mapping.items():
            if pdf_field_name in pdf_fields_data and excel_field_name in excel_data:
                field_info = pdf_fields_data[pdf_field_name]
                field_obj = field_info['obj']
                field_type = field_info['type']
                value_to_fill = excel_data[excel_field_name]

                st.write(f"Processing field '{pdf_field_name}' (Type: {field_type}) with Excel value '{value_to_fill}'")

                update_dict = pdfrw.PdfDict()

                # --- Field Type Specific Logic ---
                if field_type == pdfrw.PdfName('Btn'):  # Checkbox/Button
                    # Define values that mean "check the box"
                    check_values = ['yes', 'true', '1', 'on', 'x', 'checked']
                    normalized_value = str(value_to_fill).lower().strip()
                    
                    if normalized_value in check_values:
                        # Use the detected 'On' value, or default to 'Yes'
                        on_value = field_info.get('on_value', pdfrw.PdfName('Yes'))
                        update_dict[pdfrw.PdfName('V')] = on_value
                        update_dict[pdfrw.PdfName('AS')] = on_value
                        st.write(f"  -> Setting Checkbox state to: {on_value}")
                    else:
                        # Assume any other value means 'Off'
                        update_dict[pdfrw.PdfName('V')] = pdfrw.PdfName('Off')
                        update_dict[pdfrw.PdfName('AS')] = pdfrw.PdfName('Off')
                        st.write(f"  -> Setting Checkbox state to Off")
                
                else:  # Text field or any other type
                    encoded_value = pdfrw.PdfString.encode(str(value_to_fill))
                    update_dict[pdfrw.PdfName('V')] = encoded_value
                    update_dict[pdfrw.PdfName('DV')] = encoded_value  # Set Default Value as well
                    st.write(f"  -> Setting Text field value")

                # Apply updates to the field object
                field_obj.update(update_dict)

                # Remove appearance stream so viewer rebuilds it
                if pdfrw.PdfName('AP') in field_obj:
                    del field_obj[pdfrw.PdfName('AP')]

                # Update Kids if they exist
                if pdfrw.PdfName('Kids') in field_obj:
                    st.write(f"  -> Updating Kids for field '{pdf_field_name}'")
                    for kid in field_obj.Kids:
                        kid.update(update_dict)
                        if pdfrw.PdfName('AP') in kid:
                            del kid[pdfrw.PdfName('AP')]
                
                filled_count += 1
                processed_pdf_fields.add(pdf_field_name)
            
            elif excel_field_name not in excel_data:
                st.warning(f"Skipping field '{pdf_field_name}': Mapped Excel field '{excel_field_name}' not found in Excel data.")

        # --- Second Pass for Widgets (If using annotation method) ---
        st.write("--- Running second pass for Widget Annotations ---")
        widget_update_count = 0
        for page in pdf.pages:
            if page.Annots:
                for annot in page.Annots:
                    # Check if it's a Widget annotation with a Field Name (/T)
                    if annot.Subtype == pdfrw.PdfName('Widget') and annot.T:
                        annot_field_name = annot.T.strip('()')
                        # Check if this field was mapped AND not already processed
                        if annot_field_name in field_mapping and annot_field_name not in processed_pdf_fields:
                            excel_field_name = field_mapping[annot_field_name]
                            if excel_field_name in excel_data:
                                # We need type info for the annotation if available
                                if annot_field_name in pdf_fields_data:
                                    annot_info = pdf_fields_data[annot_field_name]
                                    annot_type = annot_info['type']
                                    value_to_fill = excel_data[excel_field_name]
                                    
                                    st.write(f"(Widget Pass) Processing field '{annot_field_name}' (Type: {annot_type})")
                                    
                                    update_dict = pdfrw.PdfDict()
                                    
                                    if annot_type == pdfrw.PdfName('Btn'):
                                        # Define values that mean "check the box"
                                        check_values = ['yes', 'true', '1', 'on', 'x', 'checked']
                                        normalized_value = str(value_to_fill).lower().strip()
                                        
                                        if normalized_value in check_values:
                                            # Use the detected 'On' value, or default to 'Yes'
                                            on_value = annot_info.get('on_value', pdfrw.PdfName('Yes'))
                                            update_dict[pdfrw.PdfName('V')] = on_value
                                            update_dict[pdfrw.PdfName('AS')] = on_value
                                            st.write(f"  -> (Widget) Setting Checkbox state to: {on_value}")
                                        else:
                                            # Assume any other value means 'Off'
                                            update_dict[pdfrw.PdfName('V')] = pdfrw.PdfName('Off')
                                            update_dict[pdfrw.PdfName('AS')] = pdfrw.PdfName('Off')
                                            st.write(f"  -> (Widget) Setting Checkbox state to Off")
                                    
                                    else:  # Text field or any other type
                                        encoded_value = pdfrw.PdfString.encode(str(value_to_fill))
                                        update_dict[pdfrw.PdfName('V')] = encoded_value
                                        update_dict[pdfrw.PdfName('DV')] = encoded_value
                                        st.write(f"  -> (Widget) Setting Text field value")
                                    
                                    # Apply update to the annotation object
                                    annot.update(update_dict)
                                    if pdfrw.PdfName('AP') in annot:
                                        del annot[pdfrw.PdfName('AP')]
                                    
                                    widget_update_count += 1
                                    processed_pdf_fields.add(annot_field_name)
                                else:
                                    st.warning(f"(Widget Pass) Type info missing for annotation '{annot_field_name}'. Cannot reliably fill.")

        st.info(f"Attempted to fill {filled_count} fields via AcroForm/main loop.")
        if widget_update_count > 0:
            st.info(f"Additionally updated {widget_update_count} fields via Widget Annotation pass.")

        # Write the modified PDF to a BytesIO object
        output_pdf_stream = io.BytesIO()
        pdfrw.PdfWriter().write(output_pdf_stream, pdf)
        output_pdf_stream.seek(0)
        return output_pdf_stream

    except Exception as e:
        st.error(f"Error filling PDF: {e}")
        import traceback
        st.error(traceback.format_exc())  # More detailed error
        return None

# --- Streamlit UI ---

st.set_page_config(layout="wide")
st.title("📄➡️📊 AI PDF Form Filler")
st.markdown("""
Upload a fillable PDF and an Excel file (.xlsx).
The app uses AI to match Excel data (Column A: Field Name, Column B: Value) to PDF fields and fills the form.
- For **Text Fields**, the value from Column B is used directly.
- For **Checkboxes**, values like `Yes`, `True`, `1`, `On`, `X`, `Checked` (case-insensitive) in Column B will check the box. Other values leave it unchecked.
""")

uploaded_pdf = st.file_uploader("1. Upload Fillable PDF", type="pdf", key="pdf_upload")
uploaded_excel = st.file_uploader("2. Upload Excel Data (.xlsx)", type="xlsx", key="excel_upload")

st.subheader("3. Process & Download")
fill_button = st.button("✨ Fill PDF using AI Matcher")

if fill_button:
    if uploaded_pdf and uploaded_excel:
        pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())
        excel_bytes_io = io.BytesIO(uploaded_excel.getvalue())

        pdf_fields_data = {}
        pdf_field_names = []
        excel_data = {}

        with st.spinner("Reading PDF fields..."):
            pdf_fields_data = get_pdf_fields(pdf_bytes_io)
            pdf_field_names = list(pdf_fields_data.keys())
            if pdf_field_names:
                st.write("Detected PDF Fields:")
                # Display field names and types for clarity
                display_df_data = [(name, str(data['type']), str(data.get('on_value', 'N/A'))) 
                                  for name, data in pdf_fields_data.items()]
                pdf_display_df = pd.DataFrame(display_df_data, 
                                             columns=['PDF Field Name', 'Type', 'Checkbox On Value'])
                st.dataframe(pdf_display_df, use_container_width=True)
            else:
                st.error("No fillable fields detected in the PDF.")
                st.stop()

        with st.spinner("Reading Excel data..."):
            excel_data = read_excel_data(excel_bytes_io)
            if excel_data:
                st.write("Detected Excel Data (Field -> Value):")
                excel_df_display = pd.DataFrame(list(excel_data.items()), 
                                               columns=['Excel Field', 'Value'])
                st.dataframe(excel_df_display, use_container_width=True)
            else:
                st.error("Could not read data from the Excel file.")
                st.stop()

        field_mapping = None
        with st.spinner(f"Matching fields with {MODEL_NAME}... This may take a moment."):
            field_mapping = match_fields_with_ai(pdf_field_names, list(excel_data.keys()))

        if field_mapping is None:
            st.error("PDF filling could not proceed due to issues in the AI matching step. Check logs above.")
            st.stop()
        elif not field_mapping:  # Empty dictionary returned
            st.warning("The AI could not confidently match any PDF fields to the Excel data provided. No fields will be filled.")
            st.stop()
        else:
            st.write("AI Field Mapping (PDF Field -> Excel Field):")
            map_df_display = pd.DataFrame(list(field_mapping.items()), 
                                         columns=['PDF Field', 'Matched Excel Field'])
            st.dataframe(map_df_display, use_container_width=True)

            with st.spinner("Filling PDF form..."):
                pdf_bytes_io.seek(0)  # Reset PDF stream before filling
                filled_pdf_stream = fill_pdf(pdf_bytes_io, field_mapping, excel_data, pdf_fields_data)

            if filled_pdf_stream:
                st.success("PDF Filled Successfully!")
                output_filename = f"filled_{uploaded_pdf.name}"
                st.download_button(
                    label="⬇️ Download Filled PDF",
                    data=filled_pdf_stream,
                    file_name=output_filename,
                    mime="application/pdf"
                )
            else:
                st.error("Failed to generate the filled PDF. Check logs for errors.")

    else:  # Files not uploaded
        if not uploaded_pdf:
            st.warning("Please upload a PDF file.")
        if not uploaded_excel:
            st.warning("Please upload an Excel file.")
