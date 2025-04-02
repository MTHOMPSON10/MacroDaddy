import os
import streamlit as st
import pandas as pd
import math
import io

# ----------------------------------------------------------
# FUNCTION DEFINITIONS
# ----------------------------------------------------------

def prob_call_waits(calls, reporting_period_minutes, average_handling_time, agents):
    """
    Calculate the probability that a call will wait.
    """
    try:
        intensity = (calls / (reporting_period_minutes * 60)) * average_handling_time
        max_occupancy_percent = intensity / agents
        A_n = 1.0
        sum_A_k = 0.0
        for k in range(agents, -1, -1):
            A_k = A_n * k / intensity
            sum_A_k += A_k
            A_n = A_k
        prob = 1 / (1 + ((1 - max_occupancy_percent) * sum_A_k))
    except Exception:
        prob = 1
    return max(0, min(prob, 1))

def service_level(calls, reporting_period_minutes, average_handling_time, service_level_time, agents):
    """
    Calculate expected service level.
    """
    try:
        intensity = (calls / (reporting_period_minutes * 60)) * average_handling_time
        prob_wait = prob_call_waits(calls, reporting_period_minutes, average_handling_time, agents)
        sl = 1 - (prob_wait * math.exp(-(agents - intensity) * service_level_time / average_handling_time))
    except Exception:
        sl = 0
    return max(0, min(sl, 1))

def occupancy(calls, reporting_period_minutes, average_handling_time, agents):
    """
    Calculate occupancy ratio.
    """
    try:
        intensity = (calls / (reporting_period_minutes * 60)) * average_handling_time
        occ = intensity / agents
    except Exception:
        occ = 0.99
    return max(0, min(occ, 0.99))

def agents_required(calls, reporting_period_minutes, average_handling_time, service_level_target, service_level_time, max_occupancy_target, shrinkage_percent):
    """
    Calculate required agents based on call centre parameters.
    """
    # Validate input ranges
    if (reporting_period_minutes < 5 or reporting_period_minutes > 1500 or
        average_handling_time < 1 or average_handling_time > 30000 or
        service_level_target < 0.00001 or service_level_target > 0.9998 or
        max_occupancy_target < 0.00001 or max_occupancy_target > 0.9998 or
        service_level_time < 1 or service_level_time > 30000 or
        shrinkage_percent < 0 or shrinkage_percent > 0.9998):
        return None

    intensity = (calls / (reporting_period_minutes * 60)) * average_handling_time
    min_agents = int(intensity)
    if min_agents < 1:
        min_agents = 1

    agents = min_agents

    # Increase agents until service level meets target
    while service_level(calls, reporting_period_minutes, average_handling_time, service_level_time, agents) < service_level_target:
        agents += 1

    # Increase agents until occupancy is within target
    while occupancy(calls, reporting_period_minutes, average_handling_time, agents) > max_occupancy_target:
        agents += 1

    # Adjust for shrinkage
    shrinkage_percent = max(0, min(shrinkage_percent, 0.99))
    required = agents / (1 - shrinkage_percent)

    if required > 600:
        st.warning("This calculator only works up to 600 agents. For higher numbers, please use the online version.")
    
    if calls == 0:
        return 0

    return math.ceil(required)

# ----------------------------------------------------------
# STREAMLIT APP INTERFACE
# ----------------------------------------------------------

st.title("Call Centre Helper Tools")

# Sidebar for tool selection
tool = st.sidebar.selectbox("Select a tool", ["Agent Calculator", "Expected Service Level", "Day Planner"])

# -------------------------
# AGENT CALCULATOR TOOL
# -------------------------
if tool == "Agent Calculator":
    st.header("Agent Calculator")
    st.markdown("Calculate the required number of agents based on call volume and performance targets.")
    
    # Inputs for Agent Calculator
    calls = st.number_input("Number of Calls", min_value=0, value=100, step=1, key="ac_calls")
    reporting_period = st.number_input("Reporting Period (minutes)", min_value=5, value=30, step=1, key="ac_reporting")
    aht = st.number_input("Average Handling Time (seconds)", min_value=1, value=360, step=1, key="ac_aht")
    service_level_target = st.slider("Service Level Target (as a decimal)", min_value=0.0, max_value=1.0, value=0.80, step=0.01, key="ac_sl")
    service_level_time = st.number_input("Service Level Time (seconds)", min_value=1, value=30, step=1, key="ac_time")
    max_occupancy_target = st.slider("Maximum Occupancy (as a decimal)", min_value=0.0, max_value=1.0, value=0.85, step=0.01, key="ac_occ")
    shrinkage = st.slider("Shrinkage (as a decimal)", min_value=0.0, max_value=1.0, value=0.17, step=0.01, key="ac_shrinkage")
    
    if st.button("Calculate Required Agents", key="ac_calc"):
        result = agents_required(calls, reporting_period, aht, service_level_target, service_level_time, max_occupancy_target, shrinkage)
        if result is None:
            st.error("One or more inputs are out of the allowed ranges. Please adjust your values.")
        else:
            st.success(f"Required Agents: {result}")

# -------------------------
# EXPECTED SERVICE LEVEL TOOL
# -------------------------
elif tool == "Expected Service Level":
    st.header("Expected Service Level")
    st.markdown("Calculate the expected service level, occupancy, and probability a call waits for a given number of agents.")
    
    # Inputs for Expected Service Level
    calls_sl = st.number_input("Number of Calls", min_value=0, value=100, step=1, key="esl_calls")
    reporting_period_sl = st.number_input("Reporting Period (minutes)", min_value=5, value=30, step=1, key="esl_reporting")
    aht_sl = st.number_input("Average Handling Time (seconds)", min_value=1, value=360, step=1, key="esl_aht")
    service_level_time_sl = st.number_input("Service Level Time (seconds)", min_value=1, value=30, step=1, key="esl_time")
    agents = st.number_input("Number of Agents", min_value=1, value=10, step=1, key="esl_agents")
    
    if st.button("Calculate Service Level", key="esl_calc"):
        sl = service_level(calls_sl, reporting_period_sl, aht_sl, service_level_time_sl, agents)
        occ = occupancy(calls_sl, reporting_period_sl, aht_sl, agents)
        prob_wait = prob_call_waits(calls_sl, reporting_period_sl, aht_sl, agents)
        st.success(f"Expected Service Level: {sl*100:.2f}%")
        st.info(f"Occupancy: {occ*100:.2f}%")
        st.info(f"Probability Call Waits: {prob_wait*100:.2f}%")

# -------------------------
# DAY PLANNER TOOL
# -------------------------
elif tool == "Day Planner":
    st.header("Day Planner")
    st.markdown("""
    Upload a CSV, XLSX, or XLSM file containing your call volumes for different time intervals.
    The file should have a column named **'Calls'** (and optionally a column for **'Time'**).
    Common parameters (such as Reporting Period, AHT, etc.) will be applied to each interval.
    """)
    
    uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx", "xlsm"], key="dp_upload")
    
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
            reporting_period_dp = st.number_input("Reporting Period (minutes)", min_value=5, value=30, step=1, key="dp_reporting")
            aht_dp = st.number_input("Average Handling Time (seconds)", min_value=1, value=360, step=1, key="dp_aht")
            service_level_target_dp = st.slider("Service Level Target (as a decimal)", min_value=0.0, max_value=1.0, value=0.80, step=0.01, key="dp_sl")
            service_level_time_dp = st.number_input("Service Level Time (seconds)", min_value=1, value=30, step=1, key="dp_time")
            max_occupancy_target_dp = st.slider("Maximum Occupancy (as a decimal)", min_value=0.0, max_value=1.0, value=0.85, step=0.01, key="dp_occ")
            shrinkage_dp = st.slider("Shrinkage (as a decimal)", min_value=0.0, max_value=1.0, value=0.17, step=0.01, key="dp_shrinkage")
            
            # Calculate required agents for each interval (each row in the file)
            df["Required Agents"] = df["Calls"].apply(
                lambda calls: agents_required(calls, reporting_period_dp, aht_dp, service_level_target_dp, service_level_time_dp, max_occupancy_target_dp, shrinkage_dp)
            )
            
            st.subheader("Day Planner Results")
            st.dataframe(df)
            
            # Write DataFrame to an in-memory buffer for download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False)
            st.download_button(
                label="Download Results as Excel",
                data=buffer.getvalue(),
                file_name="Day_Planner_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
