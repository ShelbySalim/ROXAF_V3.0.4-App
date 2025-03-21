import os
import pandas as pd
import streamlit as st
import tempfile
import zipfile
from io import BytesIO
import time  # For measuring operation duration

# Initialize Streamlit app
st.set_page_config(page_title="ROXAF - Client Stocklot Matching", layout="wide")

# Global variables
df_stocklot = None
df_client_needs = None

# Helper functions
def find_matching_column(df_columns, keywords):
    """Find a column in the DataFrame that matches any of the given keywords."""
    for col in df_columns:
        for keyword in keywords:
            if keyword.lower() in col.lower():
                return col
    return None

def group_client_needs_by_item_family(df_client_needs, client_name):
    """Group client needs by item family."""
    try:
        client_col = find_matching_column(df_client_needs.columns, ["client", "customer", "name"])
        item_family_col = find_matching_column(df_client_needs.columns, ["item family", "family", "item"])
        grammage_col = find_matching_column(df_client_needs.columns, ["grammage", "weight", "gsm"])
        laize_col = find_matching_column(df_client_needs.columns, ["laize", "width", "size"])

        if not all([client_col, item_family_col, grammage_col, laize_col]):
            st.error("Required columns not found in client needs file.")
            return None

        client_needs = df_client_needs[df_client_needs[client_col] == client_name].copy()
        if client_needs.empty:
            st.error(f"No needs found for client: {client_name}")
            return None

        client_needs.loc[:, grammage_col] = pd.to_numeric(client_needs[grammage_col], errors="coerce")
        client_needs.loc[:, laize_col] = pd.to_numeric(client_needs[laize_col], errors="coerce")
        client_needs_cleaned = client_needs.dropna(subset=[grammage_col, laize_col])

        grouped_needs = client_needs_cleaned.groupby(item_family_col).agg({
            grammage_col: ['min', 'max'],
            laize_col: ['min', 'max']
        }).reset_index()

        grouped_needs.columns = [' '.join(col).strip() for col in grouped_needs.columns.values]
        return grouped_needs
    except Exception as e:
        st.error(f"Error grouping client needs: {e}")
        return None

def filter_stocklot_for_client(df_stocklot, grouped_needs):
    """Filter stocklot data based on grouped client needs."""
    try:
        item_family_col_stocklot = find_matching_column(df_stocklot.columns, ["item family", "family", "item"])
        grammage_col_stocklot = find_matching_column(df_stocklot.columns, ["grammage", "weight", "gsm"])
        laize_col_stocklot = find_matching_column(df_stocklot.columns, ["laize", "width", "size"])

        if not all([item_family_col_stocklot, grammage_col_stocklot, laize_col_stocklot]):
            st.error("Required columns not found in stocklot file.")
            return None

        filtered_results = []
        for _, row in grouped_needs.iterrows():
            grammage_min_col = [col for col in grouped_needs.columns if 'grammage min' in col.lower()][0]
            grammage_max_col = [col for col in grouped_needs.columns if 'grammage max' in col.lower()][0]
            laize_min_col = [col for col in grouped_needs.columns if 'laize min' in col.lower()][0]
            laize_max_col = [col for col in grouped_needs.columns if 'laize max' in col.lower()][0]

            item_family = row[grouped_needs.columns[0]]
            min_grammage = row[grammage_min_col]
            max_grammage = row[grammage_max_col]
            min_laize = row[laize_min_col]
            max_laize = row[laize_max_col]

            df_filtered = df_stocklot[
                (df_stocklot[item_family_col_stocklot] == item_family) &
                (df_stocklot[grammage_col_stocklot].between(min_grammage, max_grammage)) &
                (df_stocklot[laize_col_stocklot].between(min_laize, max_laize))
            ]
            filtered_results.append(df_filtered)

        return pd.concat(filtered_results, ignore_index=True) if filtered_results else None
    except Exception as e:
        st.error(f"Error filtering stocklot: {e}")
        return None

def classify_needs_by_priority(df):
    """Classify client needs by priority."""
    try:
        priority_col = find_matching_column(df.columns, ["priority", "urgency", "importance"])
        if not priority_col:
            st.error("Priority column not found in client needs file.")
            return None

        urgent_needs = df[df[priority_col].str.lower().str.contains("urgent")]
        less_urgent_needs = df[df[priority_col].str.lower().str.contains("less urgent")]
        last_priority_needs = df[df[priority_col].str.lower().str.contains("last priority")]
        general_needs = df[~df[priority_col].str.lower().str.contains("urgent|less urgent|last priority")]

        return {
            "Urgent": urgent_needs,
            "Less Urgent": less_urgent_needs,
            "Last Priority": last_priority_needs,
            "General": general_needs,
        }
    except Exception as e:
        st.error(f"Error classifying needs by priority: {e}")
        return None

# Streamlit app
def main():
    global df_stocklot, df_client_needs

    st.title("ROXAF - Client Stocklot Matching")

    # File Upload Section
    st.header("Upload Files")
    col1, col2 = st.columns(2)  # Split into 2 columns
    with col1:
        stocklot_file = st.file_uploader("Upload Stocklot File", type=["xlsx"])
        if stocklot_file:
            df_stocklot = pd.read_excel(stocklot_file)
            st.success("Stocklot file uploaded successfully!")
    with col2:
        client_needs_file = st.file_uploader("Upload Client Needs File", type=["xlsx"])
        if client_needs_file:
            df_client_needs = pd.read_excel(client_needs_file)
            st.success("Client needs file uploaded successfully!")

    # Filtering Section
    st.header("Filtering Options")

    # Manual Filter
    st.subheader("Manual Filter")
    client_name = st.text_input("Enter client name for manual filter", max_chars=50, key="manual_filter_input")
    manual_filter = st.button("Manual Filter", key="manual_filter_btn", use_container_width=True)
    if manual_filter:
        if df_stocklot is None or df_client_needs is None:
            st.error("Please upload both files first.")
        elif not client_name:
            st.error("Please enter a client name.")
        else:
            start_time = time.time()  # Start timing
            grouped_needs = group_client_needs_by_item_family(df_client_needs, client_name)
            if grouped_needs is None:
                st.error(f"No needs found for {client_name}.")
            else:
                df_filtered = filter_stocklot_for_client(df_stocklot, grouped_needs)
                if df_filtered is None or df_filtered.empty:
                    st.error(f"No matching stocklots found for {client_name}.")
                else:
                    # Save filtered data to a temporary file
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                        df_filtered.to_excel(tmp_file.name, index=False)
                        st.success(f"Filtered data for {client_name} ready for download.")
                        st.download_button(
                            label="Download File",
                            data=open(tmp_file.name, "rb").read(),
                            file_name=f"{client_name}-ROXAF-Manual.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            end_time = time.time()  # End timing
            st.write(f"Operation duration: **{(end_time - start_time) * 1000:.2f} ms**")

    # Auto Filter
    st.subheader("Auto Filter")
    auto_filter = st.button("Auto Filter by Priority", key="auto_filter_btn", use_container_width=True)
    if auto_filter:
        if df_stocklot is None or df_client_needs is None:
            st.error("Please upload both files first.")
        else:
            start_time = time.time()  # Start timing
            classified_needs = classify_needs_by_priority(df_client_needs)
            if not classified_needs:
                st.error("Error: Priority column not found in client needs file.")
            else:
                files_to_download = []
                for priority, needs_df in classified_needs.items():
                    client_col = find_matching_column(df_client_needs.columns, ["client", "customer", "name"])
                    if not client_col:
                        st.error("Error: Client column not found in client needs file.")
                        break

                    client_names = needs_df[client_col].unique()
                    for client_name in client_names:
                        grouped_needs = group_client_needs_by_item_family(df_client_needs, client_name)
                        if grouped_needs is None:
                            continue

                        df_filtered = filter_stocklot_for_client(df_stocklot, grouped_needs)
                        if df_filtered is None or df_filtered.empty:
                            continue

                        # Save filtered data to a temporary file
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                            df_filtered.to_excel(tmp_file.name, index=False)
                            files_to_download.append((tmp_file.name, f"{client_name}-ROXAF-{priority}.xlsx"))
                            st.success(f"Filtered data for {client_name} ({priority}) ready for download.")

                # Bulk Download Option
                if files_to_download:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                        for file_path, file_name in files_to_download:
                            zip_file.write(file_path, file_name)
                    zip_buffer.seek(0)
                    st.download_button(
                        label="Download All Files as ZIP",
                        data=zip_buffer,
                        file_name="Filtered_Files.zip",
                        mime="application/zip"
                    )
            end_time = time.time()  # End timing
            st.write(f"Operation duration: **{(end_time - start_time) * 1000:.2f} ms**")

    # Check Available Selections
    st.subheader("Check Available Selections")
    if stocklot_file:
        st.write(f"Check Available Selections from list **{stocklot_file.name}**")
    check_selections = st.button("Check Available Selections", key="check_selections_btn", use_container_width=True)
    if check_selections:
        if df_stocklot is None or df_client_needs is None:
            st.error("Please upload both files first.")
        else:
            start_time = time.time()  # Start timing
            classified_needs = classify_needs_by_priority(df_client_needs)
            if not classified_needs:
                st.error("Error: Priority column not found in client needs file.")
            else:
                client_col = find_matching_column(df_client_needs.columns, ["client", "customer", "name"])
                if not client_col:
                    st.error("Error: Client column not found in client needs file.")
                else:
                    # Display client names with matching stocklots
                    st.write("### Clients with Matching Stocklots")
                    col1, col2 = st.columns(2)  # Split into 2 columns
                    for priority, needs_df in classified_needs.items():
                        client_names = needs_df[client_col].unique()
                        for idx, client_name in enumerate(client_names):
                            grouped_needs = group_client_needs_by_item_family(df_client_needs, client_name)
                            if grouped_needs is None:
                                continue

                            df_filtered = filter_stocklot_for_client(df_stocklot, grouped_needs)
                            if df_filtered is None or df_filtered.empty:
                                continue

                            # Show client name and download button
                            if idx % 2 == 0:
                                with col1:
                                    st.write(f"**{client_name}** ({priority})")
                                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                                        df_filtered.to_excel(tmp_file.name, index=False)
                                        st.download_button(
                                            label=f"Download {client_name}",
                                            data=open(tmp_file.name, "rb").read(),
                                            file_name=f"{client_name}-ROXAF-{priority}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
                            else:
                                with col2:
                                    st.write(f"**{client_name}** ({priority})")
                                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                                        df_filtered.to_excel(tmp_file.name, index=False)
                                        st.download_button(
                                            label=f"Download {client_name}",
                                            data=open(tmp_file.name, "rb").read(),
                                            file_name=f"{client_name}-ROXAF-{priority}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
            end_time = time.time()  # End timing
            st.write(f"Operation duration: **{(end_time - start_time) * 1000:.2f} ms**")

# Run the app
if __name__ == "__main__":
    main()