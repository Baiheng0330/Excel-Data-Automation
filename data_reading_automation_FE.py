import streamlit as st
import os
from backend import process_data

# Create a Streamlit app
st.title("Enter Folder and File Paths and Run Script")

# Create a text input for the folder path
df1_folder = st.text_input("Enter df1 folder path")

# Create two text inputs for file paths
df2_path = st.text_input("Enter df2 file path")
df3_path = st.text_input("Enter df3 file path")

result_path = st.text_input("Enter updated report download file path")

# Create a button to run the script
run_button = st.button("Run Script")

if run_button:
    try:
        # Call the backend function
        process_data(df1_folder, df2_path, df3_path, result_path)

        # Assume the processed Excel file is named "updated_report.xlsx"
        output_file_path = os.path.join(result_path, "updated_report.xlsx")

        # Create a download button for the processed Excel file
        with open(output_file_path, "rb") as f:
            st.download_button("Download Processed Excel File", f, file_name="updated_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.write("Script executed successfully!")

    except Exception as e:
        st.error(f"Error: {e}")
