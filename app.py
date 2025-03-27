import streamlit as st
import pandas as pd
import pdfrw  # For reading/writing PDF form fields
import json
import io
import os
# <<< CHANGE START >>>
# Import PdfName for checkbox values
from pdfrw import PdfName, PdfString, PdfDict, PdfObject
# <<< CHANGE END >>>
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

# <<< CHANGE START >>>
def get_checkbox_on_value(field):
    """Attempts to determine the 'On' value for a checkbox."""
    # Default 'On' value
    default_on_value = PdfName('Yes')

    # Checkbox fields (/Btn) often have an Appearance Dictionary (/AP)
    # with Normal (/N) states for 'Off' and the 'On' state.
    if '/AP' in field and '/N' in field['/AP']:
        ap_n = field['/AP']['/N']
        if isinstance(ap_n, PdfDict):
            # Iterate through the keys in the Normal Appearance dictionary
            for key in ap_n.keys():
                key_str = str(key)
                # The key that is not '/Off' is likely the 'On' state
                if key_str != '/Off':
                    # Return it as a PdfName
                    return PdfName(key_str.strip('/')) # Ensure it's a PdfName
    # Fallback if specific value not found
    return default_on_value

def get_pdf_fields(pdf_bytes_io):
    """Reads field names, objects, and types from a fillable PDF."""
    fields_data = {} # Store {name: {'obj': field_object, 'type': field_type, 'on_val': on_value}}
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
        if pdf.Root and '/AcroForm' in pdf.Root and '/Fields' in pdf.Root.AcroForm:
            for field in pdf.Root.AcroForm.Fields:
                field_name = field.get('/T')
                if field_name:
                    field_name = field_name.strip('()')
                    field_type = field.get('/FT') # Get field type (e.g., /Tx, /Btn)
                    on_val = None
                    if field_type == PdfName('Btn'): # Check if it's a button type
                        # Additional check: Field flags can sometimes distinguish checkboxes/radio
                        # ff_flags = field.get('/Ff', 0)
                        # is_checkbox = ff_flags & (1 << 15) # Pushbutton flag (often NOT set for checkbox)
                        # is_radio = ff_flags & (1 << 16) # Radio button flag
                        # For simplicity, treat all /Btn as potentially checkable unless clearly radio group needing separate logic
                        on_val = get_checkbox_on_value(field) # Try to find the 'On' value
                        st.write(f"Detected field: {field_name} (Type: {field_type}, OnVal: {on_val})")
                    else:
                        st.write(f"Detected field: {field_name} (Type: {field_type})")

                    fields_data[field_name] = {'obj': field, 'type': field_type, 'on_val': on_val}
            return fields_data
        else:
             st.warning("No AcroForm found in PDF Root. Trying alternative annotation search.")
             # Fall through to alternative method if AcroForm is missing/empty

    except Exception as e:
        st.error(f"Error reading PDF AcroForm fields: {e}. Trying alternative annotation search.")
        pdf_bytes_io.seek(0) # Reset stream position

    # Alternative method (less reliable for structure, but might find widgets)
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
        alt_fields = {}
        found_annot_fields = False
        for page in pdf.pages:
            if page.Annots:
                for annot in page.Annots:
                    # Check if it's a Widget annotation with a Field Name (/T)
                    if annot.Subtype == PdfName('Widget') and annot.T:
                        field_name = annot.T.strip('()')
                        # Avoid duplicates if already found via AcroForm (though we might be here because AcroForm failed)
                        if field_name not in fields_data:
                            field_type = annot.get('/FT') # Widget might have /FT
                            on_val = None
                            if field_type == PdfName('Btn'):
                                on_val = get_checkbox_on_value(annot)
                                st.write(f"(Annot) Detected field: {field_name} (Type: {field_type}, OnVal: {on_val})")
                            else:
                                st.write(f"(Annot) Detected field: {field_name} (Type: {field_type})")

                            # Store the annotation itself as the 'obj'
                            alt_fields[field_name] = {'obj': annot, 'type': field_type, 'on_val': on_val}
                            found_annot_fields = True

        if found_annot_fields:
             st.warning("Used alternative method (page annotations) to find fields. Results may vary.")
             # Combine with any AcroForm fields found earlier, preferring AcroForm if names conflict
             fields_data.update(alt_fields) # alt_fields will overwrite if duplicate names, which might be okay if AcroForm failed partially
             return fields_data
        elif not fields_data: # Only error if neither method found anything
             st.error("Could not find any form fields using standard or alternative methods.")
             return {}
        else:
             # If AcroForm fields were found but generated an error message earlier,
             # and the alternative method found nothing new, return the AcroForm fields.
             return fields_data

    except Exception as e_alt:
        st.error(f"Further error during alternative PDF field reading: {e_alt}")
        return {} # Return empty if both methods fail catastrophically

# <<< CHANGE END >>>


def read_excel_data(excel_bytes_io):
    """Reads data from the first two columns of an Excel file."""
    try:
        df = pd.read_excel(excel_bytes_io, header=None, usecols=[0, 1], engine='openpyxl')
        df = df.dropna(subset=[0])
        df[0] = df[0].astype(str)
        df[1] = df[1].fillna('').astype(str) # Ensure values are strings, handle NaNs
        data_dict = dict(zip(df[0], df[1]))
        return data_dict
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None

def match_fields_with_ai(pdf_field_names, excel_field_names):
    """Uses Mistral API to match PDF fields to Excel fields."""
    api_key = st.secrets.get("MISTRAL_API_KEY") # Use .get for safer access
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
                st.warning(f"AI proposed an invalid mapping or used invalid field names - PDF:'{pdf_key}' -> Excel:'{excel_val}'. Skipping.")

        st.success(f"{MODEL_NAME} matching complete.")
        return validated_mapping

    except Exception as e:
        st.error(f"Error calling Mistral API or processing response: {e}")
        import traceback
        st.error(traceback.format_exc()) # More detailed error for debugging
        return None

# <<< CHANGE START >>>
# Simplified update logic - moved directly into fill_pdf loop
# def update_field(field_obj, value, field_type, on_value=None):
#     """Update a field object based on its type."""
#     update_dict = PdfDict()
#     processed_value = None

#     if field_type == PdfName('Btn'): # Checkbox/Button
#         normalized_value = str(value).lower().strip()
#         check_values = ['yes', 'true', '1', 'on', 'x'] # Common values indicating 'checked'
#         if normalized_value in check_values:
#             target_state = on_value if on_value else PdfName('Yes') # Use determined or default 'On' state
#             processed_value = target_state
#             update_dict[PdfName('AS')] = target_state # Set Appearance State
#             st.write(f"  -> Setting Checkbox state to: {target_state}")
#         else:
#             target_state = PdfName('Off') # Standard 'Off' state
#             processed_value = target_state
#             update_dict[PdfName('AS')] = target_state
#             st.write(f"  -> Setting Checkbox state to: {target_state}")
#         update_dict[PdfName('V')] = processed_value # Set Value

#     elif field_type == PdfName('Tx') or field_type is None: # Text field or unknown (treat as text)
#         processed_value = PdfString.encode(str(value)) # Ensure value is string and encode
#         update_dict[PdfName('V')] = processed_value
#         update_dict[PdfName('DV')] = processed_value # Set Default Value as well for some viewers
#         st.write(f"  -> Setting Text field value")
#     else:
#         # Handle other types like /Ch (Choice) if needed later
#         st.warning(f"  -> Unsupported field type '{field_type}' for auto-filling. Trying text fill.")
#         processed_value = PdfString.encode(str(value))
#         update_dict[PdfName('V')] = processed_value
#         update_dict[PdfName('DV')] = processed_value

#     # Apply updates
#     field_obj.update(update_dict)

#     # Remove appearance stream so viewer rebuilds it
#     if PdfName('AP') in field_obj:
#         del field_obj[PdfName('AP')]

#     # Update Kids if they exist (common for radio buttons, sometimes complex fields)
#     if PdfName('Kids') in field_obj:
#         for kid in field_obj.Kids:
#             # Kid inherits properties but might need its own V/AS update
#             kid.update(update_dict)
#             if PdfName('AP') in kid:
#                 del kid[PdfName('AP')]


def fill_pdf(pdf_bytes_io, field_mapping, excel_data, pdf_fields_data):
    """Fills the PDF form fields based on the mapping and data."""
    try:
        pdf = pdfrw.PdfReader(fdata=pdf_bytes_io.getvalue())
        # Set NeedAppearances so the viewer will rebuild appearances
        if pdf.Root and pdf.Root.AcroForm:
            pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))
        else:
             st.warning("PDF does not contain an AcroForm dictionary. Field appearances might not update correctly.")
             # Create one if it doesn't exist? Risky, might break structure.
             # pdf.Root.AcroForm = PdfDict()
             # pdf.Root.AcroForm.Fields = [] # Initialize Fields array
             # pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))


        filled_count = 0
        processed_pdf_fields = set() # Keep track of fields updated via AcroForm/Annotations

        # Iterate through the AI mapping
        for pdf_field_name, excel_field_name in field_mapping.items():
            if pdf_field_name in pdf_fields_data and excel_field_name in excel_data:
                field_info = pdf_fields_data[pdf_field_name]
                field_obj = field_info['obj']
                field_type = field_info['type']
                on_value = field_info.get('on_val') # Get 'On' value if available (for checkboxes)
                value_to_fill = excel_data[excel_field_name]

                st.write(f"Processing field '{pdf_field_name}' (Type: {field_type}) with Excel value '{value_to_fill}'")

                update_dict = PdfDict()
                processed_value = None

                # --- Field Type Specific Logic ---
                if field_type == PdfName('Btn'): # Checkbox/Button
                    normalized_value = str(value_to_fill).lower().strip()
                    # Define values that mean "check the box"
                    check_values = ['yes', 'true', '1', 'on', 'x', 'checked']
                    if normalized_value in check_values:
                        # Use the detected 'On' value, or default to '/Yes'
                        target_state = on_value if on_value else PdfName('Yes')
                        processed_value = target_state
                        update_dict[PdfName('AS')] = target_state # Set Appearance State
                        st.write(f"  -> Setting Checkbox state to: {target_state}")
                    else:
                        # Assume any other value means 'Off'
                        target_state = PdfName('Off') # Standard 'Off' state
                        processed_value = target_state
                        update_dict[PdfName('AS')] = target_state
                        st.write(f"  -> Setting Checkbox state to: {target_state}")
                    update_dict[PdfName('V')] = processed_value # Set Value (/V)

                elif field_type == PdfName('Tx') or field_type is None: # Text field or unknown type
                    processed_value = PdfString.encode(str(value_to_fill)) # Ensure string and encode
                    update_dict[PdfName('V')] = processed_value
                    # Some viewers rely on /DV (Default Value) as well, setting both is safer
                    update_dict[PdfName('DV')] = processed_value
                    st.write(f"  -> Setting Text field value")

                # Add elif for /Ch (Choice fields - dropdowns, list boxes) if needed later
                # elif field_type == PdfName('Ch'):
                #    # Logic for choice fields - needs to match one of the /Opt values
                #    st.warning(f"Field '{pdf_field_name}' is a Choice field - filling as text, may not work.")
                #    processed_value = PdfString.encode(str(value_to_fill))
                #    update_dict[PdfName('V')] = processed_value

                else: # Default fallback for other known types (e.g., /Sig) or unexpected ones
                     st.warning(f"  -> Unsupported field type '{field_type}' for field '{pdf_field_name}'. Attempting to fill as text.")
                     processed_value = PdfString.encode(str(value_to_fill))
                     update_dict[PdfName('V')] = processed_value
                     update_dict[PdfName('DV')] = processed_value

                # --- Apply Updates to Field Object ---
                if field_obj is not None:
                    field_obj.update(update_dict)

                    # Remove appearance stream (/AP) so the viewer rebuilds it based on /V and /AS
                    if PdfName('AP') in field_obj:
                        del field_obj[PdfName('AP')]

                    # Update Kids if they exist (essential for some fields, esp. radio buttons grouped under one name)
                    if PdfName('Kids') in field_obj:
                         st.write(f"  -> Updating {len(field_obj.Kids)} Kid(s) for field '{pdf_field_name}'")
                         for kid in field_obj.Kids:
                             # Kids inherit properties but might need their own V/AS update matching the parent
                             # Important: Use the same update_dict derived from the parent's logic
                             kid.update(update_dict)
                             # Also remove Kid's appearance stream
                             if PdfName('AP') in kid:
                                 del kid[PdfName('AP')]

                    filled_count += 1
                    processed_pdf_fields.add(pdf_field_name)
                else:
                     st.warning(f"Field object for '{pdf_field_name}' not found during update phase.")

            elif excel_field_name not in excel_data:
                 st.warning(f"Skipping field '{pdf_field_name}': Mapped Excel field '{excel_field_name}' not found in Excel data.")
            # Don't warn if pdf_field_name not in pdf_fields_data, as it might be handled by widget pass below if using annotations

        # --- Second Pass for Widgets (If using annotation method) ---
        # This ensures widgets found via page annotations also get updated,
        # even if they weren't in the main AcroForm /Fields array.
        st.write("--- Running second pass for Widget Annotations ---")
        widget_update_count = 0
        for page in pdf.pages:
             if page.Annots:
                 for annot in page.Annots:
                     # Check if it's a Widget annotation with a Field Name (/T)
                     if annot.Subtype == PdfName('Widget') and annot.T:
                         annot_field_name = annot.T.strip('()')
                         # Check if this field was mapped AND *not already processed*
                         if annot_field_name in field_mapping and annot_field_name not in processed_pdf_fields:
                             excel_field_name = field_mapping[annot_field_name]
                             if excel_field_name in excel_data:
                                 # We need type info for the annotation if available
                                 # It *should* be in pdf_fields_data if found by get_pdf_fields
                                 if annot_field_name in pdf_fields_data:
                                     annot_info = pdf_fields_data[annot_field_name]
                                     annot_type = annot_info['type']
                                     annot_on_val = annot_info.get('on_val')
                                     value_to_fill = excel_data[excel_field_name]
                                     st.write(f"(Widget Pass) Processing field '{annot_field_name}' (Type: {annot_type})")

                                     # --- Duplicate the update logic here for annotations ---
                                     update_dict = PdfDict()
                                     processed_value = None

                                     if annot_type == PdfName('Btn'):
                                         normalized_value = str(value_to_fill).lower().strip()
                                         check_values = ['yes', 'true', '1', 'on', 'x', 'checked']
                                         if normalized_value in check_values:
                                             target_state = annot_on_val if annot_on_val else PdfName('Yes')
                                             processed_value = target_state
                                             update_dict[PdfName('AS')] = target_state
                                             st.write(f"  -> (Widget) Setting Checkbox state to: {target_state}")
                                         else:
                                             target_state = PdfName('Off')
                                             processed_value = target_state
                                             update_dict[PdfName('AS')] = target_state
                                             st.write(f"  -> (Widget) Setting Checkbox state to: {target_state}")
                                         update_dict[PdfName('V')] = processed_value

                                     elif annot_type == PdfName('Tx') or annot_type is None:
                                         processed_value = PdfString.encode(str(value_to_fill))
                                         update_dict[PdfName('V')] = processed_value
                                         update_dict[PdfName('DV')] = processed_value
                                         st.write(f"  -> (Widget) Setting Text field value")
                                     else:
                                         st.warning(f"  -> (Widget) Unsupported type '{annot_type}' for '{annot_field_name}'. Filling as text.")
                                         processed_value = PdfString.encode(str(value_to_fill))
                                         update_dict[PdfName('V')] = processed_value
                                         update_dict[PdfName('DV')] = processed_value

                                     # Apply update to the annotation object
                                     annot.update(update_dict)
                                     if PdfName('AP') in annot:
                                         del annot[PdfName('AP')]

                                     widget_update_count +=1
                                     processed_pdf_fields.add(annot_field_name) # Mark as processed
                                 else:
                                      st.warning(f"(Widget Pass) Type info missing for annotation '{annot_field_name}'. Cannot reliably fill.")
                             else:
                                  st.warning(f"(Widget Pass) Mapped Excel field '{excel_field_name}' not found for '{annot_field_name}'.")


        st.info(f"Attempted to fill {filled_count} fields via AcroForm/main loop.")
        if widget_update_count > 0:
             st.info(f"Additionally updated {widget_update_count} fields via Widget Annotation pass.")


        # Debug: Re-read values (might not reflect final viewer rendering accurately)
        # st.write("Debug: Field values after updating (inspecting PDF object):")
        # for name, data in pdf_fields_data.items():
        #      obj = data['obj']
        #      val = obj.get('/V')
        #      a_state = obj.get('/AS') # Appearance state
        #      st.write(f"Field: {name}, Type: {data['type']}, /V: {val}, /AS: {a_state}")


        # Write the modified PDF to a BytesIO object
        output_pdf_stream = io.BytesIO()
        pdfrw.PdfWriter().write(output_pdf_stream, pdf)
        output_pdf_stream.seek(0)
        return output_pdf_stream

    except Exception as e:
        st.error(f"Error filling PDF: {e}")
        import traceback
        st.error(traceback.format_exc()) # More detailed error
        return None

# <<< CHANGE END >>>


# --- Streamlit UI ---

st.set_page_config(layout="wide")
st.title("📄➡️📊 AI PDF Form Filler")
st.markdown("""
Upload a fillable PDF and an Excel file (.xlsx).
The app uses AI to match Excel data (Column A: Field Name, Column B: Value) to PDF fields and fills the form.
- For **Text Fields**, the value from Column B is used directly.
- For **Checkboxes**, values like `Yes`, `True`, `1`, `On`, `X`, `Checked` (case-insensitive) in Column B will check the box. Other values leave it unchecked.
""")

# Use session state to store results between button clicks if needed (optional for this flow)
# if 'pdf_fields_data' not in st.session_state:
#     st.session_state.pdf_fields_data = {}
# if 'excel_data' not in st.session_state:
#     st.session_state.excel_data = {}
# if 'field_mapping' not in st.session_state:
#     st.session_state.field_mapping = None
# if 'filled_pdf_stream' not in st.session_state:
#     st.session_state.filled_pdf_stream = None

uploaded_pdf = st.file_uploader("1. Upload Fillable PDF", type="pdf", key="pdf_upload")
uploaded_excel = st.file_uploader("2. Upload Excel Data (.xlsx)", type="xlsx", key="excel_upload")

st.subheader("3. Process & Download")
fill_button = st.button("✨ Fill PDF using AI Matcher")

if fill_button:
    # Reset previous results if any
    # st.session_state.filled_pdf_stream = None
    # st.session_state.field_mapping = None
    # st.session_state.pdf_fields_data = {}
    # st.session_state.excel_data = {}


    if uploaded_pdf and uploaded_excel:
        pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())
        excel_bytes_io = io.BytesIO(uploaded_excel.getvalue())

        pdf_fields_data = {}
        pdf_field_names = []
        excel_data = {}

        with st.spinner("Reading PDF fields..."):
            pdf_fields_data = get_pdf_fields(pdf_bytes_io) # Now returns dict with more info
            pdf_field_names = list(pdf_fields_data.keys())
            if pdf_field_names:
                st.write("Detected PDF Fields:")
                # Display field names and types for clarity
                display_df_data = [(name, str(data['type']), str(data.get('on_val', 'N/A'))) for name, data in pdf_fields_data.items()]
                pdf_display_df = pd.DataFrame(display_df_data, columns=['PDF Field Name', 'Type', 'Checkbox On Value'])
                st.dataframe(pdf_display_df, use_container_width=True)
            else:
                st.error("No fillable fields detected in the PDF.")
                st.stop() # Stop execution if no fields

        with st.spinner("Reading Excel data..."):
            excel_data = read_excel_data(excel_bytes_io)
            if excel_data:
                st.write("Detected Excel Data (Field -> Value):")
                excel_df_display = pd.DataFrame(list(excel_data.items()), columns=['Excel Field', 'Value'])
                st.dataframe(excel_df_display, use_container_width=True)
            else:
                st.error("Could not read data from the Excel file.")
                st.stop() # Stop execution

        field_mapping = None
        with st.spinner(f"Matching fields with {MODEL_NAME}... This may take a moment."):
            field_mapping = match_fields_with_ai(pdf_field_names, list(excel_data.keys()))

        if field_mapping is None:
             st.error("PDF filling could not proceed due to issues in the AI matching step. Check logs above.")
             st.stop()
        elif not field_mapping: # Empty dictionary returned
            st.warning("The AI could not confidently match any PDF fields to the Excel data provided. No fields will be filled.")
            st.stop()
        else:
            st.write("AI Field Mapping (PDF Field -> Excel Field):")
            map_df_display = pd.DataFrame(list(field_mapping.items()), columns=['PDF Field', 'Matched Excel Field'])
            st.dataframe(map_df_display, use_container_width=True)

            with st.spinner("Filling PDF form..."):
                pdf_bytes_io.seek(0) # Reset PDF stream before filling
                filled_pdf_stream = fill_pdf(pdf_bytes_io, field_mapping, excel_data, pdf_fields_data)

            if filled_pdf_stream:
                st.success("PDF Filled Successfully!")
                output_filename = f"filled_{uploaded_pdf.name}"
                # Store stream in session state IF you want download button to persist after reruns
                # st.session_state.filled_pdf_stream = filled_pdf_stream
                st.download_button(
                    label="⬇️ Download Filled PDF",
                    data=filled_pdf_stream, # Use the generated stream directly
                    file_name=output_filename,
                    mime="application/pdf"
                )
            else:
                st.error("Failed to generate the filled PDF. Check logs for errors.")

    else: # Files not uploaded
        if not uploaded_pdf:
            st.warning("Please upload a PDF file.")
        if not uploaded_excel:
            st.warning("Please upload an Excel file.")

# Optional: Display download button if results are in session state (useful if other interactions cause reruns)
# elif st.session_state.get('filled_pdf_stream'):
#     output_filename = f"filled_{st.session_state.get('pdf_filename', 'document.pdf')}" # Need to store filename too
#     st.download_button(
#         label="⬇️ Download Filled PDF (Previous Result)",
#         data=st.session_state.filled_pdf_stream,
#         file_name=output_filename,
#         mime="application/pdf"
#     )
