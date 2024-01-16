import streamlit as st
import os, json
from docx import Document
from docx2pdf import convert

word_path = "template.docx"
word_output_path = "invoice.docx"
companies_file_path = "companies.json"

# Load existing companies data from the file if it exists
if os.path.exists(companies_file_path):
    with open(companies_file_path, "r") as file:
        existing_companies = json.load(file)
else:
    existing_companies = {}

replacements = [
    ("Company Name", "The NU B.V", "New Company Name"),
    ("Company Address", "J.H Oortweg 21", "New Company Address"),
    ("Company Post Code", "2333CH Leiden, Netherlands", "New Company Post Code"),
    ("Charging Task", "Timc", "New Task"),
    ("Amount", "1.300,00", "New Amount"),
    ("Date Range", "29.11.2023-31.12.2023", "New Date Range")
]

def update_existing_companies(replacements_processed):
    new_company_info = {}
    for category, old_text, new_text in replacements_processed:
        if category in ["Company Name", "Company Address", "Company Post Code"]:
            new_company_info[category] = new_text

    if new_company_info:
        new_company_name = new_company_info.get("Company Name", "New Company")
        existing_company = existing_companies.get(new_company_name, {})

        if existing_company:
            # Update existing company information
            existing_company.update(new_company_info)
        else:
            # Add a new company entry
            existing_companies[new_company_name] = new_company_info

        # Save the updated companies data to the file
        with open(companies_file_path, "w") as file:
            json.dump(existing_companies, file)

def replace_text_in_word(doc, replacements):
    # Replace text in each category
    for category, old_text, new_text in replacements:
        for paragraph in doc.paragraphs:
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text)

def main():
    # Invoke the Streamlit App
    st.title("Invoice Maker")
    st.markdown(
    """
    This invoicing app is designed to streamline your invoice creation process. The template used in this app is based on a Slovenian format.

    *Note: The template used here is based on Slovenian invoicing practices, but can be applied to other countries as well.*
    """
    )

    st.sidebar.title("Existing Companies")
    selected_company = st.sidebar.selectbox("Select Company:", [""] + list(existing_companies.keys()))
    if selected_company:
        selected_company_info = existing_companies[selected_company]
        st.sidebar.write("### Company Information")
        st.sidebar.write(f"**Company Name:** {selected_company_info['Company Name']}")
        st.sidebar.write(f"**Company Address:** {selected_company_info['Company Address']}")
        st.sidebar.write(f"**Company Post Code:** {selected_company_info['Company Post Code']}")

    # User input for replacement parameters
    replacements_processed = []
    for category, old_text, _ in replacements:
        if selected_company:
            if category in ["Company Name", "Company Address", "Company Post Code"]:
                new_text = existing_companies[selected_company][category]
                replacements_processed.append((category, old_text, new_text))
        else:
            new_text = st.text_input(f"{category}:", old_text)
            replacements_processed.append((category, old_text, new_text))

    update_existing_companies(replacements_processed)
    word_output_path = st.text_input("What would you like to name this file?", word_output_path)
    apply_ddv = st.checkbox("Apply DDV (Value Added Tax)")

    # Open the Word document
    doc = Document(word_path)

    # Replace text in the Word document
    replace_text_in_word(doc, replacements_processed)

    if apply_ddv:
        # Assuming "amount" is in the replacements list
        for idx, (category, old_text, new_text) in enumerate(replacements):
            if category == "amount":
                try:
                    amount = float(new_text.replace(",", "."))  # Convert to float
                    ddv_amount = amount * 0.095  # 9.5% DDV
                    replacements[idx] = (category, old_text, f"{ddv_amount:.2f}")
                    st.success("DDV applied successfully!")
                except ValueError:
                    st.error("Invalid amount format. Please enter a valid numeric amount.")

    
    # Columns for download buttons
    col1, col2 = st.columns(2)

    # Download button for the modified Word document
    if col1.button("Download Invoice as Docx"):
        try:
            with open(word_output_path, "rb") as f:
                col1.download_button(
                    label="Download Invoice as Docx",
                    data=f.read(),
                    key="download_modified_doc",
                    file_name=os.path.basename(word_output_path),
                )
            st.success("Docx saved successfully!")
            update_existing_companies(replacements_processed=replacements_processed)
        except Exception as e:
            st.error(f"Error saving the document as PDF: {e}")

    # Option to save as PDF
    if col2.button("Download Invoice as PDF"):
        try:
            pdf_output_path = word_output_path.replace(".docx", ".pdf")
            convert(word_output_path, pdf_output_path)
            with open(pdf_output_path, "rb") as f:
                col2.download_button(
                    label="Download Invoice as PDF",
                    data=f.read(),
                    key="download_modified_pdf",
                    file_name=os.path.basename(pdf_output_path),
                )
            st.success("PDF saved successfully!")
            update_existing_companies(replacements_processed=replacements_processed)
        except Exception as e:
            st.error(f"Error saving the document as PDF: {e}")
    
     ## How to Use:
    st.markdown(
    """
    
        ## How to Use:

        - **User Inputs:** Start by filling out the necessary information such as the recipient, address, task, amount, company name, and date range.

        - **Existing Companies:** If you've invoiced the same company before, you can select an existing company from the dropdown menu to auto-fill the details.

        - **Download Options:** Once your invoice is ready, you can download it as a Word document or PDF.

        - **DDV (VAT):** Optionally, you can enable the DDV (VAT) checkbox to apply a 9.5% tax to the total amount.

        Feel free to explore the features and enjoy a smoother invoicing experience!

            2024. All rights reserved. Created by TimC.
        """
    )

if __name__ == "__main__":
    main()