import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
from qc_logger import process_qc_submission

def qc_radar_app():
    st.markdown('<div class="custom-heading">Knock knock! Your HDVI MVR wants to be validated 👀</div>', unsafe_allow_html=True)
    st.write("Upload an MVR Excel file to validate.")

    def run_qc(df: pd.DataFrame):
        today = date.today()
        df.columns = [str(col).strip() for col in df.columns]
        df.replace(regex={r"(?i)^\s*$": pd.NA, r"(?i)^N/A$": pd.NA}, inplace=True)

        required_fields = [
            "First Name", "Last Name", "Date of Birth", "Age",
            "Years of Experience", "Hire Date", "Years of Tenure",
            "License State", "CDL Number", "CDL Type", "Expiration Date"
        ]
        allowed_cdl_types = {"A", "B", "C"}

        results = []
        tag_counts = {"valid": 0, "partial": 0, "fail": 0}
        total_deductions = 0

        for idx, row in df.iterrows():
            issues = []
            score = 100

            def deduct(points, reason):
                nonlocal score
                score -= points
                issues.append(reason)

            for field in required_fields:
                if pd.isna(row.get(field)) or str(row.get(field)).strip() == "":
                    deduct(10, f"{field} missing")

            try:
                dob = pd.to_datetime(row["Date of Birth"], errors="coerce").date()
                if dob >= today:
                    deduct(15, "DOB is in the future")
            except:
                dob = None

            try:
                age = int(row["Age"])
                if dob:
                    expected_age = (today - dob).days // 365
                    if abs(expected_age - age) > 1:
                        deduct(15, f"Age mismatch (expected ~{expected_age}, got {age})")
            except:
                pass

            try:
                exp = float(row["Years of Experience"])
                if 'age' in locals() and exp > age - 18:
                    deduct(10, f"Experience too high for age {age}")
            except:
                pass

            try:
                hire = pd.to_datetime(row["Hire Date"], errors="coerce").date()
                if dob and hire < dob.replace(year=dob.year + 18):
                    deduct(20, "Hire before legal age (18)")
            except:
                hire = None

            try:
                tenure = float(row["Years of Tenure"])
                if hire:
                    expected_tenure = (today - hire).days / 365.25
                    if abs(expected_tenure - tenure) > 1:
                        deduct(5, f"Tenure mismatch (expected ~{expected_tenure:.2f}, got {tenure})")
            except:
                pass

            try:
                exp_date = pd.to_datetime(row["Expiration Date"], errors="coerce").date()
                if exp_date < today:
                    deduct(20, f"License expired on {exp_date}")
            except:
                pass

            cdl_type = str(row.get("CDL Type", "")).strip().upper()
            if cdl_type and cdl_type not in allowed_cdl_types:
                deduct(5, f"Unexpected CDL Type '{cdl_type}'")

            cdl_num = str(row.get("CDL Number", "")).strip()
            if cdl_num and not cdl_num.replace("-", "").isalnum():
                deduct(5, "CDL Number not alphanumeric")

            if score >= 90:
                tag = "✅ Valid"
                tag_counts["valid"] += 1
            elif score >= 75:
                tag = "⚠️ Partial"
                tag_counts["partial"] += 1
            else:
                tag = "❌ High Risk"
                tag_counts["fail"] += 1

            results.append({"QC Tag": tag, "QC Issues": "; ".join(issues) if issues else "OK"})
            total_deductions += (100 - score)

        total_possible = 100 * len(df)
        confidence = round((1 - (total_deductions / total_possible)) * 100, 2)

        qc_df = df.copy().reset_index(drop=True)
        qc_df = pd.concat([qc_df, pd.DataFrame(results)], axis=1)

        return qc_df, confidence, tag_counts

    uploaded = st.file_uploader("📂 Upload Excel File", type=["xlsx"])
    if uploaded:
        try:
            raw_df = pd.read_excel(uploaded, sheet_name="hdvi output", header=None)
            raw_df.columns = raw_df.iloc[1]
            df = raw_df[2:].reset_index(drop=True)
            df.columns = [str(c).strip() for c in df.columns]

            if "MVR Received" not in df.columns:
                st.error("🚫 'MVR Received' column is missing. Please include it to proceed.")
                st.stop()

            client_driver_count = st.number_input("🧑‍💼 Enter total number of drivers from client list:", min_value=1)
            expected_true_count = st.number_input("✅ Enter number of drivers expected to have MVRs (TRUE):", min_value=0, max_value=client_driver_count)

            if client_driver_count and expected_true_count is not None:
                df["MVR Received"] = df["MVR Received"].astype(str).str.upper().str.strip()
                actual_true = (df["MVR Received"] == "TRUE").sum()
                actual_false = (df["MVR Received"] == "FALSE").sum()
                total_labeled = actual_true + actual_false

                st.markdown("### 🔍 MVR Received Breakdown")
                st.write(f"✅ Marked TRUE: {actual_true}")
                st.write(f"❌ Marked FALSE: {actual_false}")
                st.write(f"📄 Total Labeled: {total_labeled}")
                st.write(f"🧑‍💼 Declared Client Count: {int(client_driver_count)}")
                st.write(f"🎯 Expected TRUEs: {int(expected_true_count)}")

                if total_labeled != client_driver_count:
                    st.error("🚫 Mismatch: TRUE + FALSE must equal client driver count. Please check your file.")
                    st.stop()

                if actual_true != expected_true_count:
                    st.error("❗ TRUE count mismatch. File cannot be processed until corrected.")
                    st.stop()

                st.success("✅ MVR count validation passed. Proceeding to QC...")

                qc_df, confidence, counts = run_qc(df)

                st.success(f"✅ QC completed! Confidence Score: **{confidence}%**")
                st.dataframe(qc_df[["QC Tag", "QC Issues"]], use_container_width=True)

                buffer = BytesIO()
                qc_df.to_excel(buffer, index=False, engine="openpyxl")
                buffer.seek(0)

                st.download_button(
                    label="📥 Download QC Report",
                    data=buffer,
                    file_name="qc_results_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.markdown("### 📊 QC Summary")
                st.write(f"✅ Valid Rows: {counts['valid']}")
                st.write(f"⚠️ Partial Rows: {counts['partial']}")
                st.write(f"❌ High Risk Rows: {counts['fail']}")

                # ✅ Submit Task Logging
                st.markdown("### 📝 Submit This Task")
                task_type = st.selectbox("Select Task Type", ["HDVI MVR", "Alltrans MVR", "IFTA"])

                if "username" in st.session_state:
                    if st.button("📌 Submit QC Result"):
                        result = process_qc_submission(
                            user=st.session_state["username"],
                            task_type=task_type,
                            confidence=confidence
                        )
                        st.success(f"🎯 Task {result['TaskID']} logged for **{result['UserID']}** — Accuracy: **{confidence}%**")
                else:
                    st.warning("🔐 Please login to submit this result.")

        except Exception as e:
            st.error(f"An error occurred while processing your file: {e}")
