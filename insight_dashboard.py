import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta, time

def insight_dashboard_app():
    st.title("📊 Oogway Insights Dashboard")

    try:
        df = pd.read_csv("task_data.csv", parse_dates=["SubmissionTime"])
    except FileNotFoundError:
        st.warning("No task data found yet. Users need to submit via QC Radar.")
        return

    if "username" not in st.session_state or "role" not in st.session_state:
        st.error("User not authenticated.")
        return

    current_user = st.session_state["username"]
    role = st.session_state["role"]
    st.markdown(f"### 👤 Welcome, **{current_user}** ({role.upper()})")

    # Filter view based on role
    if role.lower() == "admin":
        st.markdown("#### 🧮 All Users Leaderboard")

        leaderboard = (
            df.groupby("UserID")
              .agg(Tasks=("TaskID", "count"),
                   Avg_Accuracy=("QC_Passed", "mean"),
                   Last_Submission=("SubmissionTime", "max"))
              .reset_index()
              .sort_values("Avg_Accuracy", ascending=False)
        )
        leaderboard["Avg_Accuracy"] = (leaderboard["Avg_Accuracy"] * 100).round(2)
        st.dataframe(leaderboard, use_container_width=True)

        fig = px.bar(
            leaderboard, x="UserID", y="Avg_Accuracy", text="Tasks",
            title="🔍 Accuracy by User", color="Avg_Accuracy",
            color_continuous_scale="greens"
        )
        st.plotly_chart(fig, use_container_width=True)

        selected_user = st.selectbox("👥 Select user for detailed view", df["UserID"].unique())
        user_data = df[df["UserID"] == selected_user]
    else:
        st.markdown("#### 🧠 Your QC Performance")
        user_data = df[df["UserID"] == current_user]

    if user_data.empty:
        st.info("No task submissions found for this user.")
        return

    # 🧾 Performance Summary
    total_tasks = user_data.shape[0]
    avg_accuracy = round(user_data["QC_Passed"].mean() * 100, 2)
    last_time = user_data["SubmissionTime"].max().strftime("%b %d, %Y %I:%M %p")
    st.metric("📦 Total Tasks", total_tasks)
    st.metric("🎯 Avg Accuracy", f"{avg_accuracy}%")
    st.metric("🕒 Last Submission", last_time)

    # 📅 Shift-Day Mapping (6:30 PM to 3:30 AM)
    def map_qc_shift_day(ts):
        ts = ts.tz_localize('Asia/Kolkata') if ts.tzinfo is None else ts
        if ts.time() >= time(18, 30):  # After 6:30 PM
            return ts.strftime("%A")
        elif ts.time() <= time(3, 30):  # Before 3:30 AM — previous day
            return (ts - pd.Timedelta(days=1)).strftime("%A")
        return None

    user_data["QC_Shift_Day"] = user_data["SubmissionTime"].apply(map_qc_shift_day)
    recent_cutoff = datetime.now() - timedelta(days=5)
    last_5_days = user_data[user_data["SubmissionTime"] >= recent_cutoff]
    last_5_days = last_5_days.dropna(subset=["QC_Shift_Day"])

    # 🔍 Shift-based Accuracy Breakdown
    st.markdown("#### 📆 Accuracy by QC Shift Day (6:30 PM – 3:30 AM)")
    if last_5_days.empty:
        st.info("No submissions in the past 5 shift days.")
    else:
        shift_summary = (
            last_5_days.groupby("QC_Shift_Day")
                       .agg(Tasks=("TaskID", "count"), Accuracy=("QC_Passed", "mean"))
                       .reset_index()
        )
        day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        shift_summary["Accuracy"] = (shift_summary["Accuracy"] * 100).round(2)
        shift_summary["QC_Shift_Day"] = pd.Categorical(shift_summary["QC_Shift_Day"], categories=day_order, ordered=True)
        shift_summary = shift_summary.sort_values("QC_Shift_Day")

        st.dataframe(shift_summary, use_container_width=True)
        fig_day = px.bar(
            shift_summary, x="QC_Shift_Day", y="Accuracy", text="Tasks",
            title="🎯 Accuracy by QC Day",
            color="Accuracy", color_continuous_scale="blues"
        )
        st.plotly_chart(fig_day, use_container_width=True)

    # 📈 Accuracy Over Time (Timeline)
    st.markdown("#### 🪄 Full Task Timeline")
    timeline = user_data.sort_values("SubmissionTime", ascending=True)
    fig2 = px.line(
        timeline, x="SubmissionTime", y="QC_Passed", markers=True,
        title="Trend of QC Outcomes Over Time"
    )
    fig2.update_traces(line=dict(color="#00BFFF"))
    st.plotly_chart(fig2, use_container_width=True)
