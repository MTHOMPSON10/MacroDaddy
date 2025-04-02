import os
import streamlit as st
import pandas as pd
import io
import math

# --- A Simplified agents_required function ---
# Replace this logic with your full calculation as needed.
def agents_required(calls, reporting_period_minutes, average_handling_time, 
                    service_level_target, service_level_time, max_occupancy_target, shrinkage):
    # Calculate call intensity
    intensity = (calls / (reporting_period_minutes * 60)) * average_handling_time
    # Start with a minimum number of agents (at least 1)
    agents = max(1, int(intensity))
    # For demonstration, we simply adjust for shrinkage:
    required = agents / (1 - shrinkage)
    return math.ceil(required)

# --- Day Planner Section ---
st.header("Day Planner")
st.markdown("""
Upload a CSV or Excel file containing your call volumes for different time intervals.
The file should have a column named **'Calls'** (and optionally a column for **'Time'**).
Common parameters (such as Reporting Period, AHT, etc.) will be applied to each interval.
""")

# File uploader accepts CSV, XLSX, and XLSM files
uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx", "xlsm"])

if uploaded_file:
    filename = uploaded_file.name
    extension = os.path.splitext(filename)[1].lower()
    
    try:
        if extension == ".xlsx":
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif extension == ".xlsm":
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif extension == ".csv":
            df = pd.read_csv(uploaded_file)
        else:
            st.error("Please upload a valid .xlsx, .xlsm, or .csv file.")
            st.stop()
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()
    
    if "Calls" not in df.columns:
        st.error("The uploaded file must contain a column named 'Calls'.")
    else:
        # Common parameters for each interval
        reporting_period = st.number_input("Reporting Period (minutes)", min_value=5, value=30, step=1, key="dp_reporting")
        aht = st.number_input("Average Handling Time (seconds)", min_value=1, value=360, step=1, key="dp_aht")
        service_level_target = st.slider("Service Level Target (as a decimal)", min_value=0.0, max_value=1.0, value=0.80, step=0.01, key="dp_sl")
        service_level_time = st.number_input("Service Level Time (seconds)", min_value=1, value=30, step=1, key="dp_time")
        max_occupancy_target = st.slider("Maximum Occupancy (as a decimal)", min_value=0.0, max_value=1.0, value=0.85, step=0.01, key="dp_occ")
        shrinkage = st.slider("Shrinkage (as a decimal)", min_value=0.0, max_value=1.0, value=0.17, step=0.01, key="dp_shrinkage")
        
        # Calculate Required Agents for each row based on the "Calls" column
        df["Required Agents"] = df["Calls"].apply(
            lambda calls: agents_required(calls, reporting_period, aht, service_level_target, service_level_time, max_occupancy_target, shrinkage)
        )
        
        st.subheader("Day Planner Results")
        st.dataframe(df)
        
        # Write DataFrame to an in-memory buffer
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        # Download button to retrieve the results
        st.download_button(
            label="Download Results as Excel",
            data=buffer.getvalue(),
            file_name="Day_Planner_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
