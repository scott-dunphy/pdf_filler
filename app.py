import streamlit as st
import pandas as pd
import PyPDF2
import openai
import os

# --------------------------------------------------------------------------
# 1) Configure your OpenAI API key
# --------------------------------------------------------------------------
# Option 1: You have OPENAI_API_KEY in your environment
# openai.api_key = os.getenv("OPENAI_API_KEY")

# Option 2: Use Streamlit secrets (uncomment if you're storing the key in streamlit secrets)
# openai.api_key = st.secrets["openai_api_key"]

# --------------------------------------------------------------------------
# 2) OpenAI LLM Matching
# --------------------------------------------------------------------------
def openai_match_field(pdf_field: str, excel_columns: list, model_name: str = "gpt-4o-mini") -> str:
    """
    Uses OpenAI ChatCompletion to determine which Excel column best matches a given PDF field.
    pdf_field: The name of the PDF form field (e.g., "Company Name")
    excel_columns: List of column names in the Excel (e.g., ["Business Name", "Address", "Contact Person"])
    model_name: The OpenAI model to use (defaults to GPT-3.5-turbo)
    Returns the single best matching column name, or "" if none is confidently matched.
    """

    # Build a concise conversation to guide the model.
    # We provide a system message describing its role and constraints.
    # Then a user message that includes the PDF field and the possible Excel columns.
    system_message = {
        "role": "system",
        "content": (
            "You are a helpful assistant specialized in matching a PDF form field name "
            "to the best Excel column name. Always return exactly one column name from the provided list. "
            "If no suitable match, return 'NONE'."
        ),
    }

    user_message = {
        "role": "user",
        "content": (
            f"PDF field: {pdf_field}\n"
            f"Possible columns: {excel_columns}\n\n"
            "Return the column name that best matches the PDF field. If you have no match, return 'NONE'."
        ),
    }

    # Call the ChatCompletion endpoint
    response = openai.ChatCompletion.create(
        model=model_name,
        messages=[system_message, user_message],
        temperature=0.0,  # Lower temperature for more deterministic output
        max_tokens=50,
    )

    # The response content is presumably the single column name (or 'NONE').
    answer = response["choices"][0]["message"]["content"].strip()

    # Validate the response:
    # If the model returns a column name in the list, we'll use it.
    # If it returns 'NONE' or something else, we default to "".
    if answer in excel_columns:
        return answer
    else:
        return ""

# --------------------------------------------------------------------------
# 3) PDF Form Field Extraction & Filling
# --------------------------------------------------------------------------
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

# --------------------------------------------------------------------------
# 4) The Main Streamlit App
# --------------------------------------------------------------------------
def main():
    st.title("AI Fillable PDF Filler (OpenAI version)")

    # Let user input (or hide) their OpenAI key
    # (If you're not using environment vars or st.secrets)
    user_openai_key = st.text_input("Enter your OpenAI API key (Optional)", type="password")
    if user_openai_key:
        openai.api_key = user_openai_key

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

        # For simplicity, assume we only use the first row
        excel_row = df.iloc[0].to_dict()
        excel_columns = list(excel_row.keys())

        st.write("**Step 1: OpenAI Field Matching**")
        st.write("Below are the PDF fields and the AI's guess for the best Excel column. Adjust if needed.")

        # Use OpenAI to propose a match for each PDF field
        proposed_map = {}
        user_map = {}

        with st.spinner("Matching fields..."):
            for pdf_field in pdf_fields:
                matched_col = openai_match_field(pdf_field, excel_columns)
                proposed_map[pdf_field] = matched_col

        # Let the user override each match
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
