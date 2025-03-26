import streamlit as st
import pandas as pd
import PyPDF2

# -----------------------------------------------------------------------------
# 1. GPT-4o-mini Matching Logic (Stub)
# -----------------------------------------------------------------------------
def llm_match_fields(pdf_field, excel_columns):
    """
    Example function for GPT-4o-mini to match a single PDF field name to the best
    Excel column. 
    *This function is a placeholder—replace with your actual call to GPT-4o-mini.* 

    pdf_field: str - A single PDF field name (e.g., 'Company Name')
    excel_columns: List[str] - A list of Excel column names (e.g., ['Business Name', 'Address', 'Contact Person'])

    Returns a single string: the name of the most likely matching column, or "" if no confident match was found.
    """

    # PSEUDO-CODE:

    # 1. Build a prompt for your GPT-4o-mini model.
    prompt = f"""
    You are a specialized model for matching PDF form fields to Excel column names.
    The PDF field is: "{pdf_field}"
    Possible Excel columns: {excel_columns}

    Return the single column name that best corresponds to the PDF field.
    If you have no good match, return an empty string.
    """

    # 2. Send it to your GPT-4o-mini. For example, if you have a local inference function:
    #    (Replace `my_local_gpt_model_inference` with your actual function or API call.)
    # response = my_local_gpt_model_inference(prompt)

    # 3. Parse the response to extract the single column name.
    #    This might involve some regex or direct text parsing, depending on how GPT-4o-mini responds.
    # For demonstration, we return a placeholder answer:
    # (In practice, you'd interpret the LLM’s actual text. E.g., the response might be: "Business Name".)
    response = "Business Name"  # <--- Replace with parsed response from GPT-4o-mini.

    # 4. Validate the response
    if response in excel_columns:
        return response
    else:
        # If the response isn't recognized, return empty or None
        return ""

# -----------------------------------------------------------------------------
# 2. PDF Form Extraction and Filling
# -----------------------------------------------------------------------------
def extract_pdf_fields(pdf_reader):
    """
    Extract form field names from a fillable PDF using PyPDF2.
    Returns a list of field names.
    """
    fields = set()
    if "/AcroForm" in pdf_reader.trailer["/Root"]:
        form = pdf_reader.trailer["/Root"]["/AcroForm"]
        if "/Fields" in form:
            for field in form["/Fields"]:
                field_obj = field.get_object()
                if "/T" in field_obj:
                    fields.add(field_obj["/T"])
    return list(fields)

def fill_pdf_fields(pdf_reader, data_map):
    """
    Create a new PDF in memory with fields filled in according to data_map.
    data_map should be {pdf_field_name: value_to_fill}.
    Returns a PyPDF2.PdfWriter object.
    """
    pdf_writer = PyPDF2.PdfWriter()
    pdf_writer.clone_document_from_reader(pdf_reader)

    for page_num in range(len(pdf_writer.pages)):
        page = pdf_writer.pages[page_num]
        pdf_writer.update_page_form_field_values(page, data_map)
    return pdf_writer

# -----------------------------------------------------------------------------
# 3. The Main Streamlit App
# -----------------------------------------------------------------------------
def main():
    st.title("AI Fillable PDF Filler (GPT-4o-mini version)")

    # 1. Upload fillable PDF
    pdf_file = st.file_uploader("Upload the fillable PDF", type=["pdf"])

    # 2. Upload Excel file
    excel_file = st.file_uploader("Upload the Excel file", type=["xlsx", "xls"])

    if pdf_file and excel_file:
        # Read the PDF
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        pdf_fields = extract_pdf_fields(pdf_reader)

        # Read the Excel
        df = pd.read_excel(excel_file)
        st.write("Excel Preview:")
        st.dataframe(df.head())

        # For simplicity, assume we only use the first row of the Excel
        excel_row = df.iloc[0].to_dict()
        excel_columns = list(excel_row.keys())

        st.write("**Step 1: GPT-4o-mini Field Matching**")
        st.write("Below are the PDF fields and GPT-4o-mini's guesses for the best Excel column. Adjust as needed.")

        proposed_map = {}
        user_map = {}

        # Use GPT-4o-mini to propose a match for each PDF field
        for pdf_field in pdf_fields:
            matched_col = llm_match_fields(pdf_field, excel_columns)
            proposed_map[pdf_field] = matched_col

        # Let the user verify / adjust each match
        for pdf_field in pdf_fields:
            col_options = [""] + excel_columns
            default_index = col_options.index(proposed_map[pdf_field]) if proposed_map[pdf_field] in col_options else 0
            selected_col = st.selectbox(
                f"Match PDF field '{pdf_field}' to column:",
                col_options,
                index=default_index
            )
            user_map[pdf_field] = selected_col

        if st.button("Fill PDF"):
            # Build the final fill map for PDF form fields
            final_data_map = {}
            for pdf_field, col_name in user_map.items():
                if col_name in excel_row:
                    final_data_map[pdf_field] = str(excel_row[col_name])
                else:
                    final_data_map[pdf_field] = ""

            # Fill the PDF
            filled_pdf_writer = fill_pdf_fields(pdf_reader, final_data_map)

            # Output as a downloadable PDF
            output_pdf_bytes = filled_pdf_writer.output(dest="S").read()

            st.download_button(
                label="Download Filled PDF",
                data=output_pdf_bytes,
                file_name="filled_form.pdf",
                mime="application/pdf"
            )

if __name__ == "__main__":
    main()
