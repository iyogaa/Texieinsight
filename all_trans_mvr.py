import streamlit as st
import pandas as pd
import io
from datetime import datetime


def all_trans_mvr_app(get_valid_column, names_match):
    st.markdown('<div class="custom-heading">Alltrans Excel Creation</div>', unsafe_allow_html=True)
    
    def driver_matching_app():
        st.header("Upload Files")
        col1, col2 = st.columns(2)
        with col1:
            driver_file = st.file_uploader("DRIVER LIST", type=["xlsx"])
        with col2:
            output_file = st.file_uploader("OUTPUT FILE", type=["xlsx"])
        if not driver_file or not output_file:
            st.info("Please upload both files to proceed")
            return
        st.header("Configuration...")
        driver_skip = st.number_input("Rows to skip in DRIVER file", min_value=0, value=0)
        output_skip = 3  # Fixed as requested
        try:
            xls = pd.ExcelFile(output_file)
            sheet_names = xls.sheet_names
            sheet = st.selectbox("Select sheet to process", sheet_names, 
                                 index=sheet_names.index("All Trans") if "All Trans" in sheet_names else 0)
            drivers = pd.read_excel(driver_file, skiprows=driver_skip)
            output = pd.read_excel(output_file, sheet_name=sheet, skiprows=output_skip)
            st.header("Column Mapping Running...")
            st.info("Map columns between files. The tool will try to auto-detect columns.")
            driver_name_col = get_valid_column(drivers, "driver names", ['name', 'driver name', 'full name'])
            hire_date_col = get_valid_column(drivers, "hire dates", ['hire date', 'date of hire', 'doh'], False)
            dob_col = get_valid_column(drivers, "date of birth", ['dob', 'date of birth', 'birth date'], False)
            license_col = get_valid_column(drivers, "license state", ['license state', 'lic state', 'state'], False)
            st.write(f"Detected Driver Name Column: `{driver_name_col}`")
            if hire_date_col: st.write(f"Detected Hire Date Column: `{hire_date_col}`")
            if dob_col: st.write(f"Detected Date of Birth Column: `{dob_col}`")
            if license_col: st.write(f"Detected License State Column: `{license_col}`")
            output_name_col = get_valid_column(output, "driver names", ['Name of Driver', 'Driver Name', 'Name'])
            output_dob_col = get_valid_column(output, "date of birth", ['DOB', 'Date of Birth'], False) or "DOB"
            output_license_col = get_valid_column(output, "license state", ['Lic State', 'License State', 'State'], False) or "Lic State"
            output_notes_col = get_valid_column(output, "notes", ['Notes', 'Remarks', 'Comments']) or "Notes"
            output_hire_col = get_valid_column(output, "hire date", ['DOH', 'Hire Date', 'Date of Hire'], False) or "DOH"
            st.write(f"Detected Driver Name Column: `{output_name_col}`")
            st.write(f"Detected DOB Column: `{output_dob_col}`")
            st.write(f"Detected License State Column: `{output_license_col}`")
            st.write(f"Detected Notes Column: `{output_notes_col}`")
            st.write(f"Detected Hire Date Column: `{output_hire_col}`")
            for col in [output_dob_col, output_license_col, output_notes_col, output_hire_col]:
                if col not in output.columns:
                    output[col] = ""
            if st.button("Process File", use_container_width=True):
                with st.spinner("Matching names..."):
                    match_count = 0
                    total_original = len(output)
                    matched_driver_indices = set()
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    for idx, row in output.iterrows():
                        output_name = row[output_name_col]
                        matched = False
                        for driver_idx, driver_row in drivers.iterrows():
                            if driver_idx in matched_driver_indices:
                                continue
                            if names_match(output_name, driver_row[driver_name_col]):
                                matched_driver_indices.add(driver_idx)
                                output.at[idx, output_notes_col] = "MATCH FOUND"
                                if hire_date_col:
                                    output.at[idx, output_hire_col] = driver_row[hire_date_col]
                                if dob_col:
                                    output.at[idx, output_dob_col] = driver_row[dob_col]
                                if license_col:
                                    output.at[idx, output_license_col] = driver_row[license_col]
                                match_count += 1
                                matched = True
                                break
                        progress = (idx + 1) / total_original
                        progress_bar.progress(progress)
                        status_text.text(f"Processed {idx + 1}/{total_original} records")
                    new_rows = []
                    for driver_idx, driver_row in drivers.iterrows():
                        if driver_idx not in matched_driver_indices:
                            new_row = {col: "" for col in output.columns}
                            new_row[output_name_col] = driver_row[driver_name_col]
                            if hire_date_col:
                                new_row[output_hire_col] = driver_row[hire_date_col]
                            if dob_col:
                                new_row[output_dob_col] = driver_row[dob_col]
                            if license_col:
                                new_row[output_license_col] = driver_row[license_col]
                            new_row[output_notes_col] = "MISSING MVR"
                            new_rows.append(new_row)
                    if new_rows:
                        new_rows_df = pd.DataFrame(new_rows)
                        output = pd.concat([output, new_rows_df], ignore_index=True)
                        added_count = len(new_rows)
                    else:
                        added_count = 0
                    timestamp = datetime.now().strftime("%m%d%Y")
                    result_filename = f"Driver_Matching_Result_{timestamp}.xlsx"
                    output_bytes = io.BytesIO()
                    with pd.ExcelWriter(output_bytes, engine='openpyxl') as writer:
                        output.to_excel(writer, sheet_name=sheet, index=False)
                        for other_sheet in sheet_names:
                            if other_sheet != sheet:
                                pd.read_excel(output_file, sheet_name=other_sheet).to_excel(writer, sheet_name=other_sheet, index=False)
                    output_bytes.seek(0)
                    total_final = len(output)
                    st.success("Matching complete!")
                    st.subheader("📊 Results Summary")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Original Records", total_original)
                    col2.metric("Matched Records", match_count)
                    col3.metric("Added Records", added_count)
                    st.write("**Data Transferred:**")
                    if hire_date_col:
                        st.write(f"- Hire Dates: {match_count + added_count}")
                    if dob_col:
                        st.write(f"- Birth Dates: {match_count + added_count}")
                    if license_col:
                        st.write(f"- License States: {match_count + added_count}")
                    st.download_button(
                        label="Download Excel",
                        data=output_bytes,
                        file_name=result_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.subheader("Preview of Processed Data")
                    st.dataframe(output.head(10))
        except Exception as e:
            st.error(f"⚠️ Error occurred: {str(e)}")
            st.exception(e)
    driver_matching_app()