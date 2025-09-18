import streamlit as st
import os, shutil, threading, traceback
import streamlit as st
import os, traceback
from dotenv import load_dotenv
import shutil, threading, time
# Import workflow utilities
from workflow import (
    get_sheets_to_process,
    process_workflow,
    company_df_dict,
    model_filters,
    brand_short_names,
    filter_quarters,
    filter_company_models,
    add_grand_total_row,
    preprocess_royal_enfield,
    add_fy_mean_next_to_q4,
    add_growth_metrics,
    replace_large_numbers,
    dfs_to_ppt
)
# ---------------------
# Load Environment Variables
# ---------------------
load_dotenv()
API_KEY = os.getenv("API_KEY")

if not API_KEY:
    st.error("⚠️ API_KEY not found in environment. Please set it in .env")
    st.stop()
######################################################################################################3333

#########################################################################################################
# ---------------------
# Session state for login
# ---------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "output_ppt" not in st.session_state:
    st.session_state.output_ppt = None
if "values_choice" not in st.session_state:
    st.session_state.values_choice = None



# ---------------------
# Cleanup
# ---------------------
def cleanup_folders():
    """Clear all files inside input/ and pptd/ folders, keep folders."""
    try:
        for folder in ["input", "pptd"]:
            if os.path.exists(folder):
                for filename in os.listdir(folder):
                    file_path = os.path.join(folder, filename)
                    try:
                        if os.path.isfile(file_path):
                            os.remove(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                    except Exception as e:
                        print(f"⚠️ Could not delete {file_path}: {e}")
        print("✅ Cleanup completed (files cleared, folders kept)")
    except Exception as e:
        print("⚠️ Cleanup failed:", e)


# ---------------------
# Login
# ---------------------
def login():
    st.title("🔑 Login")
    api_key = st.text_input("Enter API Key", type="password")

    if st.button("Login"):
        if api_key == API_KEY:
            st.session_state.authenticated = True
            st.success("✅ Login successful! Redirecting...")
            st.rerun()
        else:
            st.error("❌ Invalid API Key")


# ---------------------
# Upload & Generate
# ---------------------
def upload_page():
    st.title("📤 Upload Excel to Generate PPT")

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

    values_choice = st.radio(
        "Select Data Type:",
        ["Exports", "Sales"],
        horizontal=True
    )

    if uploaded_file is not None and st.button("Generate PPT"):
        try:
            os.makedirs("input", exist_ok=True)
            os.makedirs("pptd", exist_ok=True)

            file_path = os.path.join("input", uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            output_ppt = f"pptd/ppt-{values_choice}.pptx"

            process_workflow(
                    file_path=file_path,
                    sheets_to_process=get_sheets_to_process(file_path),
                    company_df_dict=company_df_dict,
                    model_filters=model_filters,
                    brand_short_names=brand_short_names,
                    filter_quarters=filter_quarters,
                    filter_company_models=filter_company_models,
                    add_grand_total_row=add_grand_total_row,
                    preprocess_royal_enfield=preprocess_royal_enfield,
                    add_fy_mean_next_to_q4=add_fy_mean_next_to_q4,
                    add_growth_metrics=add_growth_metrics,
                    replace_large_numbers=replace_large_numbers,
                    dfs_to_ppt=dfs_to_ppt,
                    values=values_choice,
                    output_ppt=output_ppt
                )


            st.session_state.output_ppt = output_ppt
            st.session_state.values_choice = values_choice

            st.success(f"🎉 PPT for {values_choice} generated successfully!")
            st.rerun()  # Refresh to show download button

        except Exception:
            st.error("❌ Error while generating PPT")
            st.code(traceback.format_exc())

    # Show download button only after generation
    if st.session_state.output_ppt:
        with open(st.session_state.output_ppt, "rb") as f:
            if st.download_button(
                label=f"⬇️ Download PPT ({st.session_state.values_choice})",
                data=f,
                file_name=f"ppt-{st.session_state.values_choice}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ):
                # Schedule cleanup in 20 seconds
                threading.Timer(20, cleanup_folders).start()
                st.info("🧹 Temporary files will be cleaned in 20 seconds...")
                # Reset state after download
                st.session_state.output_ppt = None
                st.session_state.values_choice = None
                st.rerun()

    if st.button("Logout"):
        st.session_state.authenticated = False
        st.session_state.output_ppt = None
        st.session_state.values_choice = None
        st.info("🔒 Logged out successfully")
        st.rerun()


# ---------------------
# Main
# ---------------------
if not st.session_state.authenticated:
    login()
else:
    upload_page()
