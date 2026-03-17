import streamlit as st
import subprocess
import os

st.set_page_config(page_title="MACL vs RAMIS Tool", layout="centered")

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
        try:
            # --------------------------------------------------
            # CREATE OUTPUT FOLDER
            # --------------------------------------------------
            os.makedirs("output", exist_ok=True)

            # --------------------------------------------------
            # SAVE FILES (IMPORTANT FIX)
            # --------------------------------------------------
            with open("ramis.xlsx", "wb") as f:
                f.write(winter_file.read())

            with open("macl_master.xlsx", "wb") as f:
                f.write(macl_file.read())

            with open("connecting.xlsx", "wb") as f:
                f.write(connect_file.read())

            # --------------------------------------------------
            # RUN YOUR SCRIPTS
            # --------------------------------------------------
            subprocess.run(["python", "step1_match_macl_ramis.py"], check=True)
            subprocess.run(["python", "step2_ramis_clean.py"], check=True)
            subprocess.run(["python", "step3_purple_fill.py"], check=True)
            subprocess.run(["python", "step4_reconcile.py"], check=True)

            # --------------------------------------------------
            # OUTPUT FILE
            # --------------------------------------------------
            output_file = "output/FINAL_OUTPUT.xlsx"

            if os.path.exists(output_file):
                with open(output_file, "rb") as f:
                    st.success("Process completed successfully")
                    st.download_button(
                        "Download Final Output",
                        f,
                        file_name="FINAL_OUTPUT.xlsx"
                    )
            else:
                st.error("Output file not generated")

        except Exception as e:
            st.error(f"Error occurred: {e}")
