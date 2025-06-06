from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import openai
import os

# Load environment variables
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

# Streamlit App Configuration
st.set_page_config(page_title="Excel Chat with OpenAI", layout="wide")
st.header("📊 Chat with Your Excel File")

# File uploader for Excel files
uploaded_file = st.file_uploader("Upload an Excel file (XLSX/XLS)...", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read all sheets into a dict of DataFrames
    try:
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
    else:
        st.subheader("🗂️ Table Schemas")
        sheet_dfs = excel_data  # { sheet_name: DataFrame }

        # Display schema for each sheet
        for sheet_name, df in sheet_dfs.items():
            st.markdown(f"**Sheet: {sheet_name}**")
            schema_df = pd.DataFrame({
                "Column Name": df.columns,
                "Data Type": [str(dtype) for dtype in df.dtypes]
            })
            st.table(schema_df)

        st.markdown("---")
        st.subheader("💬 Ask a Question About Your Data")
        question = st.text_input("Enter your question here:")

        if question:
            with st.spinner("Getting answer from OpenAI…"):
                # Build a context string including sheet names, schemas, and a preview of data
                context_parts = []
                for sheet_name, df in sheet_dfs.items():
                    # List of columns
                    cols = ", ".join(list(df.columns))
                    # First 5 rows as CSV-like text
                    preview = df.head(5).to_csv(index=False)
                    part = (
                        f"Sheet Name: {sheet_name}\n"
                        f"Columns: {cols}\n"
                        f"Data Preview (first 5 rows):\n{preview}\n"
                    )
                    context_parts.append(part)
                full_context = "\n".join(context_parts)

                # Construct the prompt for the model
                prompt = (
                    "You are a helpful data assistant. Use the following Excel data to answer the user's question.\n\n"
                    f"{full_context}\n"
                    f"Question: {question}\n"
                    "Provide a clear, concise answer based on the data."
                )

                # Call OpenAI ChatCompletion
                try:
                    response = openai.ChatCompletion.create(
                        model="gpt-3.5-turbo",
                        messages=[
                            {"role": "system", "content": "You are a helpful data assistant."},
                            {"role": "user", "content": prompt}
                        ],
                        temperature=0.2,
                    )
                    answer = response.choices[0].message.content.strip()
                except Exception as e:
                    st.error(f"OpenAI API error: {e}")
                    answer = None

            if answer:
                st.subheader("🤖 Answer")
                st.write(answer)
