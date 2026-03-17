import streamlit as st
import subprocess
import os

st.title("MACL vs RAMIS Automation Tool")

st.write("Upload required Excel files")

# Upload files
winter_file = st.file_uploader("Upload Winter Schedule", type=["xlsx"])
macl_file = st.file_uploader("Upload MASTER MACL", type=["xlsx"])
connect_file = st.file_uploader("Upload Connecting Flights", type=["xlsx"])

if st.button("Run Process"):

    if not winter_file or not macl_file or not connect_file:
        st.error("Please upload all required files")
    else:
        os.makedirs("input", exist_ok=True)
        os.makedirs("output", exist_ok=True)

        # Save uploaded files
        with open("input/Winter Schedule 2025.xlsx", "wb") as f:
            f.write(winter_file.read())

        with open("input/MASTER MACL WINTER vs RAMIS.xlsx", "wb") as f:
            f.write(macl_file.read())

        with open("input/Connecting Flight Plans.xlsx", "wb") as f:
            f.write(connect_file.read())

        # Run scripts
        subprocess.run(["python", "step1_match_macl_ramis.py"])
        subprocess.run(["python", "step2_ramis_clean.py"])
        subprocess.run(["python", "step3_purple_fill.py"])
        subprocess.run(["python", "step4_reconcile.py"])

        # Output
        output_file = "output/FINAL_OUTPUT.xlsx"

        if os.path.exists(output_file):
            with open(output_file, "rb") as f:
                st.success("Process completed")
                st.download_button("Download Final Output", f, file_name="FINAL_OUTPUT.xlsx")
        else:
            st.error("Something went wrong")
