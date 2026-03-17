import streamlit as st
import subprocess
import os
import sys

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
        st.stop()

    try:
        # CREATE FOLDERS
        os.makedirs("input", exist_ok=True)
        os.makedirs("output", exist_ok=True)

        # SAVE FILES
        ramis_path = os.path.abspath("input/ramis.xlsx")
        macl_path = os.path.abspath("input/macl_master.xlsx")
        conn_path = os.path.abspath("input/connecting.xlsx")

        with open(ramis_path, "wb") as f:
            f.write(winter_file.getbuffer())

        with open(macl_path, "wb") as f:
            f.write(macl_file.getbuffer())

        with open(conn_path, "wb") as f:
            f.write(connect_file.getbuffer())

        st.success("Files uploaded and saved successfully")

        # ENV VARIABLES
        env = os.environ.copy()
        env["RAMIS_FILE"] = ramis_path
        env["MACL_FILE"] = macl_path
        env["CONNECT_FILE"] = conn_path

        # RUN FUNCTION
        def run_script(script_name):
            result = subprocess.run(
                [sys.executable, script_name],
                capture_output=True,
                text=True,
                env=env
            )

            if result.returncode != 0:
                st.error(f"{script_name} FAILED")
                st.code(result.stderr)
                st.stop()
            else:
                st.success(f"{script_name} completed")

        # RUN STEPS
        run_script("step1_match_macl_ramis.py")
        run_script("step2_ramis_clean.py")
        run_script("step3_purple_fill.py")
        run_script("step4_reconcile.py")

        # DOWNLOAD OUTPUT
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
        st.error("Critical error occurred")
        st.code(str(e))
