# ==========================================================
# MACL vs RAMIS AUTOMATION TOOL (STREAMLIT VERSION)
# ==========================================================

import streamlit as st
import subprocess
import os
import sys

# ----------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------
st.set_page_config(
    page_title="MACL vs RAMIS Tool",
    layout="centered"
)

st.title("MACL vs RAMIS Automation Tool")
st.write("Upload required Excel files to process reconciliation")

# ----------------------------------------------------------
# FILE UPLOADS
# ----------------------------------------------------------
winter_file = st.file_uploader("Upload Winter Schedule (RAMIS)", type=["xlsx"])
macl_file = st.file_uploader("Upload MASTER MACL", type=["xlsx"])
connect_file = st.file_uploader("Upload Connecting Flights", type=["xlsx"])

# ----------------------------------------------------------
# MAIN PROCESS BUTTON
# ----------------------------------------------------------
if st.button("Run Process"):

    # VALIDATION
    if not winter_file or not macl_file or not connect_file:
        st.error("Please upload all required files before running the process")
        st.stop()

    try:
        # --------------------------------------------------
        # CREATE REQUIRED DIRECTORIES
        # --------------------------------------------------
        INPUT_DIR = "input"
        OUTPUT_DIR = "output"

        os.makedirs(INPUT_DIR, exist_ok=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # --------------------------------------------------
        # SAVE UPLOADED FILES
        # --------------------------------------------------
        ramis_path = os.path.abspath(os.path.join(INPUT_DIR, "ramis.xlsx"))
        macl_path = os.path.abspath(os.path.join(INPUT_DIR, "macl_master.xlsx"))
        conn_path = os.path.abspath(os.path.join(INPUT_DIR, "connecting.xlsx"))

        with open(ramis_path, "wb") as f:
            f.write(winter_file.getbuffer())

        with open(macl_path, "wb") as f:
            f.write(macl_file.getbuffer())

        with open(conn_path, "wb") as f:
            f.write(connect_file.getbuffer())

        st.success("Files uploaded and saved successfully")

        # --------------------------------------------------
        # ENVIRONMENT VARIABLES (PASS TO SCRIPTS)
        # --------------------------------------------------
        env = os.environ.copy()
        env["RAMIS_FILE"] = ramis_path
        env["MACL_FILE"] = macl_path
        env["CONNECT_FILE"] = conn_path

        # --------------------------------------------------
        # SCRIPT EXECUTION FUNCTION
        # --------------------------------------------------
        def run_script(script_name):
            st.info(f"Running {script_name}...")

            result = subprocess.run(
                [sys.executable, script_name],
                capture_output=True,
                text=True,
                env=env
            )

            if result.returncode != 0:
                st.error(f"{script_name} FAILED ❌")
                st.code(result.stderr)
                st.stop()
            else:
                st.success(f"{script_name} completed ✅")

        # --------------------------------------------------
        # EXECUTION PIPELINE (ORDER MATTERS)
        # --------------------------------------------------
        run_script("step1_match_macl_ramis.py")
        run_script("step2_ramis_clean.py")
        run_script("step3_purple_fill.py")
        run_script("step4_reconcile.py")

        # --------------------------------------------------
        # OUTPUT DOWNLOAD
        # --------------------------------------------------
        output_file = os.path.join(OUTPUT_DIR, "FINAL_OUTPUT.xlsx")

        if os.path.exists(output_file):
            with open(output_file, "rb") as f:
                st.success("Process completed successfully 🎯")

                st.download_button(
                    label="Download Final Output",
                    data=f,
                    file_name="FINAL_OUTPUT.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("Output file not generated. Check script outputs.")

    except Exception as e:
        st.error("Critical error occurred ❌")
        st.code(str(e))
