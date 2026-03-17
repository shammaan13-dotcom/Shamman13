import streamlit as st
import match_script
import purple_script

st.title("MACL / RAMIS Processing Tool")

macl_file = st.file_uploader("Upload MASTER MACL File", type=["xlsx"])
ramis_file = st.file_uploader("Upload RAMIS File", type=["xlsx"])

if macl_file and ramis_file:

    st.write("Processing...")

    # STEP 1 → MATCH
    step1_output = match_script.process(macl_file, ramis_file)

    # STEP 2 → PURPLE UPDATE
    final_output = purple_script.process(step1_output, ramis_file)

    st.success("Completed")

    st.download_button(
        label="Download Final Updated File",
        data=final_output,
        file_name="FINAL_MACL_OUTPUT.xlsx"
    )
