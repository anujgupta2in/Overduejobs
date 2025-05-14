import streamlit as st
import pandas as pd
from datetime import datetime
import utils
import io
import os
import re

st.set_page_config(
    page_title="Job Status Analyzer",
    page_icon="📊",
    layout="wide"
)

# Add custom CSS
st.markdown("""
    <style>
    .stProgress > div > div > div > div {
        background-color: #F63366;
    }
    .stDownloadButton button {
        background-color: #F63366;
        color: white;
        border-radius: 5px;
        padding: 0.5rem 1rem;
    }
    div[data-testid="stDataFrame"] div[role="cell"] {
        font-family: monospace;
    }
    div[data-testid="stDataFrame"] div[role="columnheader"] {
        background-color: #1F4E78;
        color: white;
        font-weight: bold;
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid #E0E0E0;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# Title and description
st.title("📊 Job Status Analyzer")
st.markdown("""
    Upload CSV files containing job status information to analyze:
    - Total job counts per vessel
    - New job counts
    - Generate formatted Excel reports
""")

# File uploader
uploaded_files = st.file_uploader(
    "Upload CSV files",
    type=['csv'],
    accept_multiple_files=True,
    help="Select one or more CSV files containing job status information"
)

# Initialize analysis_results at the global level
analysis_results = None

if uploaded_files:
    # Process files
    progress_bar = st.progress(0)
    status_text = st.empty()

    summary_data = []
    all_file_data = []  # Store raw CSV data for overdue analysis
    
    for i, file in enumerate(uploaded_files):
        status_text.text(f"Processing {file.name}...")
        # Process summary data
        file_summary = utils.process_csv_file(file)
        summary_data.append(file_summary)
        
        # Also store the raw CSV data for overdue job analysis
        try:
            # Reset file pointer to beginning of file
            file.seek(0)
            
            # Read the CSV file into a DataFrame and add file name to track source
            file_df = pd.read_csv(file)
            file_df['_source_file'] = file.name
            all_file_data.append(file_df)
        except Exception as e:
            st.warning(f"Error reading {file.name} for detailed analysis: {str(e)}")
        
        progress_bar.progress((i + 1) / len(uploaded_files))

    # Create DataFrame with summary data
    df = pd.DataFrame(summary_data)
    
    # Combine all raw file data for overdue analysis
    if all_file_data:
        try:
            combined_file_data = pd.concat(all_file_data, ignore_index=True)
            # Perform overdue analysis on all files in one go
            analysis_results = utils.analyze_overdue_jobs(combined_file_data)
        except Exception as e:
            st.error(f"Error analyzing overdue jobs: {str(e)}")
            # Initialize empty analysis results if analysis fails
            analysis_results = {
                'file_results': [],
                'overdue_jobs_count': 0,
                'overdue_jobs_percentage': 0,
                'critical_overdue_jobs_count': 0,
                'critical_overdue_jobs_percentage': 0,
                'total_jobs': 0,
                'overdue_jobs': pd.DataFrame(),
                'critical_overdue_jobs': pd.DataFrame()
            }

    # Convert date strings to datetime for filtering
    df['Date Extracted from File Name'] = pd.to_datetime(
        df['Date Extracted from File Name'],
        format='%d-%m-%Y',
        errors='coerce'
    )

    # Filters
    st.subheader("📌 Filters")
    col1, col2 = st.columns(2)

    with col1:
        vessel_filter = st.multiselect(
            "Filter by Vessel Name",
            options=sorted(df['Vessel Name'].unique()),
            help="Select one or more vessels to filter the data"
        )

    with col2:
        min_date = df['Date Extracted from File Name'].min()
        max_date = df['Date Extracted from File Name'].max()
        date_range = st.date_input(
            "Select Date Range",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date()
        )

    # Apply filters
    filtered_df = df.copy()
    if vessel_filter:
        filtered_df = filtered_df[filtered_df['Vessel Name'].isin(vessel_filter)]
    if len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['Date Extracted from File Name'].dt.date >= start_date) &
            (filtered_df['Date Extracted from File Name'].dt.date <= end_date)
        ]

    # Summary Statistics
    st.subheader("📈 Summary Statistics")

    # Overall metrics for files and vessels
    file_vessel_metrics = st.columns(2)
    file_vessel_metrics[0].metric("Total Files", len(filtered_df))
    file_vessel_metrics[1].metric("Total Vessels", filtered_df['Vessel Name'].nunique())
    
    # Job status metrics - display file breakdown with overdue metrics
    st.markdown("### Job Status Overview (Per File)")
    
    # Create a base table for job stats per file
    job_status_table = pd.DataFrame(filtered_df[['File Name', 'Vessel Name', 'Total Count of Jobs', 'New Job Count']])
    
    # Try to match file-level overdue analysis with the files
    if analysis_results and 'file_results' in analysis_results and analysis_results['file_results']:
        # Create mapping of file names to their analysis results
        file_analysis_map = {}
        for file_result in analysis_results['file_results']:
            file_name = file_result['file_name']
            file_analysis_map[file_name] = {
                'Overdue Jobs': file_result['overdue_jobs_count'],
                'Critical Overdue': file_result['critical_overdue_jobs_count'],
                'Overdue %': f"{file_result['overdue_jobs_percentage']}%"
            }
        
        # Extract just the filename part for easier matching
        # This helps match "path/to/file.csv" with just "file.csv"
        simplified_map = {}
        for full_path in file_analysis_map:
            # Get just the filename without path
            simple_name = os.path.basename(full_path)
            simplified_map[simple_name] = file_analysis_map[full_path]
            # Also store with the full path as key
            simplified_map[full_path] = file_analysis_map[full_path]
        
        # Add overdue metrics to the table by matching file names
        overdue_jobs = []
        critical_overdue = []
        overdue_pct = []
        
        for _, row in job_status_table.iterrows():
            file_name = row['File Name']
            simple_name = os.path.basename(file_name)
            
            # Try exact match first
            if file_name in simplified_map:
                overdue_jobs.append(simplified_map[file_name]['Overdue Jobs'])
                critical_overdue.append(simplified_map[file_name]['Critical Overdue'])
                overdue_pct.append(simplified_map[file_name]['Overdue %'])
            # Try with just the filename
            elif simple_name in simplified_map:
                overdue_jobs.append(simplified_map[simple_name]['Overdue Jobs'])
                critical_overdue.append(simplified_map[simple_name]['Critical Overdue'])
                overdue_pct.append(simplified_map[simple_name]['Overdue %'])
            else:
                # Try partial matching as a last resort
                matched = False
                for analysis_file in file_analysis_map:
                    # Check if analysis file contains this file name or vice versa
                    if file_name in analysis_file or analysis_file in file_name:
                        overdue_jobs.append(file_analysis_map[analysis_file]['Overdue Jobs'])
                        critical_overdue.append(file_analysis_map[analysis_file]['Critical Overdue'])
                        overdue_pct.append(file_analysis_map[analysis_file]['Overdue %'])
                        matched = True
                        break
                
                if not matched:
                    # Nothing matched, put placeholder values
                    overdue_jobs.append("N/A")
                    critical_overdue.append("N/A")
                    overdue_pct.append("N/A")
        
        # Add the overdue metrics to the table
        job_status_table['Overdue Jobs'] = overdue_jobs
        job_status_table['Critical Overdue'] = critical_overdue
        job_status_table['Overdue %'] = overdue_pct
    else:
        # Add placeholders if no overdue analysis is available
        job_status_table['Overdue Jobs'] = "N/A"
        job_status_table['Critical Overdue'] = "N/A"
        job_status_table['Overdue %'] = "N/A"
    
    # Display the table
    st.dataframe(
        job_status_table,
        use_container_width=True,
        hide_index=True
    )
    
    # Show the overdue jobs analysis section if analysis results exist
    st.subheader("🔍 Overdue Jobs Analysis")
    
    if not analysis_results or not analysis_results.get('file_results'):
        st.warning("""
            No overdue job analysis available. Make sure your CSV files include:
            - 'Calculated Due Date' column
            - 'Job Status' column with 'pending' or 'in progress on board' values
            - Optional criticality indicator column
        """)
    else:
        # Display overall summary metrics
        st.markdown("### Overall Summary")
        overdue_metrics = st.columns(5)
        
        overdue_metrics[0].metric(
            "Total Jobs", 
            analysis_results['total_jobs']
        )
        
        # Calculate new jobs in the detailed file (if available)
        new_status_jobs = 0
        if 'overdue_jobs' in analysis_results and not analysis_results['overdue_jobs'].empty:
            if 'Job Status' in analysis_results['overdue_jobs'].columns:
                new_status_jobs = (analysis_results['overdue_jobs']['Job Status'].str.strip().str.lower() == 'new').sum()
        
        overdue_metrics[1].metric(
            "New Status Jobs", 
            new_status_jobs
        )
        
        overdue_metrics[2].metric(
            "Total Overdue", 
            analysis_results['overdue_jobs_count']
        )
        
        overdue_metrics[3].metric(
            "Total Critical", 
            analysis_results['critical_overdue_jobs_count']
        )
        
        overdue_metrics[4].metric(
            "Overall Overdue %", 
            f"{analysis_results['overdue_jobs_percentage']}%"
        )
        
        # Show overdue jobs chart
        st.plotly_chart(
            utils.create_overdue_jobs_chart(
                analysis_results['overdue_jobs_count'],
                analysis_results['critical_overdue_jobs_count']
            ),
            use_container_width=True
        )
        
        # Display detailed file-level metrics in separate expander
        with st.expander("📊 Detailed Overdue Analysis by File"):
            # Create a DataFrame with the file-level metrics
            if analysis_results['file_results']:
                file_metrics = []
                for file_result in analysis_results['file_results']:
                    file_metrics.append({
                        'File Name': file_result['file_name'],
                        'Total Jobs': file_result['total_jobs'],
                        'Overdue Jobs': file_result['overdue_jobs_count'],
                        'Critical Overdue': file_result['critical_overdue_jobs_count'],
                        'Overdue %': f"{file_result['overdue_jobs_percentage']}%",
                        'Critical %': f"{file_result['critical_overdue_jobs_percentage']}%"
                    })
                
                # Display file metrics as a table
                st.dataframe(
                    pd.DataFrame(file_metrics),
                    use_container_width=True,
                    hide_index=True
                )
        
        # Display overdue jobs tables in expanders
        if not analysis_results['overdue_jobs'].empty:
            with st.expander("📋 View Overdue Jobs Details"):
                st.dataframe(
                    analysis_results['overdue_jobs'],
                    use_container_width=True,
                    hide_index=True
                )
        
        if not analysis_results['critical_overdue_jobs'].empty:
            with st.expander("⚠️ View Critical Overdue Jobs Details"):
                st.dataframe(
                    analysis_results['critical_overdue_jobs'],
                    use_container_width=True,
                    hide_index=True
                )

    # Data Visualizations
    st.subheader("📊 Data Visualizations")

    # Create tabs for different visualizations
    tab1, tab2 = st.tabs([
        "📊 Job Distribution", 
        "📈 Timeline Analysis"
    ])

    with tab1:
        st.plotly_chart(
            utils.create_vessel_job_distribution_chart(filtered_df, analysis_results),
            use_container_width=True
        )

    with tab2:
        st.plotly_chart(
            utils.create_jobs_timeline_chart(filtered_df, analysis_results),
            use_container_width=True
        )

    # Per-vessel detailed breakdown with expanders
    st.subheader("📊 Per Vessel File Breakdown")

    # Format the date column to show only the date
    filtered_df_display = filtered_df.copy()
    filtered_df_display['Date Extracted from File Name'] = filtered_df_display['Date Extracted from File Name'].dt.strftime('%d-%m-%Y')
    
    # Create mappings for overdue job information if available
    file_overdue_map = {}
    file_critical_map = {}
    file_overdue_pct_map = {}
    
    if analysis_results and 'file_results' in analysis_results and analysis_results['file_results']:
        for file_result in analysis_results['file_results']:
            file_name = file_result['file_name']
            # Store both exact name and basename for flexible matching
            file_overdue_map[file_name] = file_result['overdue_jobs_count']
            file_critical_map[file_name] = file_result['critical_overdue_jobs_count']
            file_overdue_pct_map[file_name] = f"{file_result['overdue_jobs_percentage']}%"
            
            # Also store with basename as key
            basename = os.path.basename(file_name)
            file_overdue_map[basename] = file_result['overdue_jobs_count']
            file_critical_map[basename] = file_result['critical_overdue_jobs_count']
            file_overdue_pct_map[basename] = f"{file_result['overdue_jobs_percentage']}%"

    # Group by vessel
    for vessel in sorted(filtered_df['Vessel Name'].unique()):
        vessel_data = filtered_df_display[filtered_df_display['Vessel Name'] == vessel]
        
        # Add overdue metrics to vessel_data if available
        if analysis_results and 'file_results' in analysis_results and analysis_results['file_results']:
            # Prepare lists for the overdue data
            overdue_jobs = []
            critical_overdue = []
            overdue_pct = []
            
            # Extract overdue metrics for each file
            for _, row in vessel_data.iterrows():
                file_name = row['File Name']
                
                # Try exact match first
                if file_name in file_overdue_map:
                    overdue_jobs.append(file_overdue_map[file_name])
                    critical_overdue.append(file_critical_map[file_name])
                    overdue_pct.append(file_overdue_pct_map[file_name])
                else:
                    # Try with basename
                    basename = os.path.basename(file_name)
                    if basename in file_overdue_map:
                        overdue_jobs.append(file_overdue_map[basename])
                        critical_overdue.append(file_critical_map[basename])
                        overdue_pct.append(file_overdue_pct_map[basename])
                    else:
                        # No match found, use placeholder
                        overdue_jobs.append("N/A")
                        critical_overdue.append("N/A")
                        overdue_pct.append("N/A")
            
            # Create a copy of vessel_data with overdue metrics
            vessel_data_with_overdue = vessel_data.copy()
            vessel_data_with_overdue['Overdue Jobs'] = overdue_jobs
            vessel_data_with_overdue['Critical Overdue'] = critical_overdue
            vessel_data_with_overdue['Overdue %'] = overdue_pct
            
            # Calculate vessel-level totals for display in the expander header
            vessel_total_jobs = vessel_data['Total Count of Jobs'].sum()
            vessel_new_jobs = vessel_data['New Job Count'].sum()
            
            # Calculate vessel-level overdue totals (excluding N/A entries)
            vessel_overdue_total = sum([x for x in overdue_jobs if isinstance(x, (int, float))])
            vessel_critical_total = sum([x for x in critical_overdue if isinstance(x, (int, float))])
            
            # Calculate vessel-level overdue percentage
            if vessel_total_jobs > 0:
                vessel_overdue_pct = f"{round((vessel_overdue_total / vessel_total_jobs) * 100, 1)}%"
            else:
                vessel_overdue_pct = "0.0%"
            
            # Create expander with overdue metrics
            with st.expander(f"🚢 {vessel} - {len(vessel_data)} files"):
                # Vessel total metrics with overdue information
                st.markdown(
                    f"**Total Jobs: {vessel_total_jobs}** | "
                    f"**New Jobs: {vessel_new_jobs}** | "
                    f"**Overdue Jobs: {vessel_overdue_total}** | "
                    f"**Critical Overdue: {vessel_critical_total}** | "
                    f"**Overdue %: {vessel_overdue_pct}**"
                )
                
                # Individual file details with overdue information
                st.dataframe(
                    vessel_data_with_overdue[['Date Extracted from File Name', 'File Name', 
                                            'Total Count of Jobs', 'New Job Count',
                                            'Overdue Jobs', 'Critical Overdue', 'Overdue %']]
                    .sort_values('Date Extracted from File Name', ascending=False),
                    use_container_width=True,
                    hide_index=True
                )
        else:
            # Create expander for each vessel (without overdue data)
            with st.expander(f"🚢 {vessel} - {len(vessel_data)} files"):
                # Vessel total metrics
                st.markdown(f"**Total Jobs: {vessel_data['Total Count of Jobs'].sum()}** | "
                           f"**New Jobs: {vessel_data['New Job Count'].sum()}**")
    
                # Individual file details
                st.dataframe(
                    vessel_data[['Date Extracted from File Name', 'File Name', 
                               'Total Count of Jobs', 'New Job Count']]
                    .sort_values('Date Extracted from File Name', ascending=False),
                    use_container_width=True,
                    hide_index=True
                )

    # Display full detailed results
    st.subheader("📋 Detailed Results")
    
    # Add overdue metrics to detailed results if available
    if analysis_results and 'file_results' in analysis_results and analysis_results['file_results']:
        # Create a copy of the filtered dataframe for display
        detailed_df = filtered_df_display.copy()
        
        # Add columns for overdue metrics
        detailed_df['Overdue Jobs'] = "N/A"
        detailed_df['Critical Overdue'] = "N/A"
        detailed_df['Overdue %'] = "N/A"
        
        # Populate overdue metrics for each file
        for i, row in detailed_df.iterrows():
            file_name = row['File Name']
            
            # Try exact match first
            if file_name in file_overdue_map:
                detailed_df.at[i, 'Overdue Jobs'] = file_overdue_map[file_name]
                detailed_df.at[i, 'Critical Overdue'] = file_critical_map[file_name]
                detailed_df.at[i, 'Overdue %'] = file_overdue_pct_map[file_name]
            else:
                # Try with basename
                basename = os.path.basename(file_name)
                if basename in file_overdue_map:
                    detailed_df.at[i, 'Overdue Jobs'] = file_overdue_map[basename]
                    detailed_df.at[i, 'Critical Overdue'] = file_critical_map[basename]
                    detailed_df.at[i, 'Overdue %'] = file_overdue_pct_map[basename]
                else:
                    # Try partial matching as last resort
                    for key in file_overdue_map:
                        if key in file_name or file_name in key:
                            detailed_df.at[i, 'Overdue Jobs'] = file_overdue_map[key]
                            detailed_df.at[i, 'Critical Overdue'] = file_critical_map[key]
                            detailed_df.at[i, 'Overdue %'] = file_overdue_pct_map[key]
                            break
        
        # Display detailed results with overdue metrics
        st.dataframe(
            detailed_df,
            use_container_width=True,
            hide_index=True
        )
    else:
        # Reorder the columns to put Date first, followed by other columns
        column_order = [
            'Date Extracted from File Name',
            'File Name', 
            'Vessel Name', 
            'Total Count of Jobs', 
            'New Job Count'
        ]
        filtered_df_display = filtered_df_display[column_order]
        
        # Display without overdue metrics
        st.dataframe(
            filtered_df_display,
            use_container_width=True,
            hide_index=True
        )

    # Download button for Excel report
    if st.button("📥 Generate Excel Report"):
        # Prepare file-level overdue data to include in the report if available
        file_level_overdue_data = None
        if analysis_results and 'file_results' in analysis_results and analysis_results['file_results']:
            file_metrics = []
            for file_result in analysis_results['file_results']:
                # Extract date from filename using regex
                date_match = re.search(r'\b(\d{2})(\d{2})(\d{4})\b', file_result['file_name'])
                formatted_date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}" if date_match else "Unknown"
                
                file_metrics.append({
                    'Date Extracted from File Name': formatted_date,
                    'File Name': file_result['file_name'],
                    'Total Jobs': file_result['total_jobs'],
                    'Overdue Jobs': file_result['overdue_jobs_count'],
                    'Critical Overdue': file_result['critical_overdue_jobs_count'],
                    'Overdue %': f"{file_result['overdue_jobs_percentage']}%",
                    'Critical %': f"{file_result['critical_overdue_jobs_percentage']}%"
                })
            file_level_overdue_data = pd.DataFrame(file_metrics)
        
        # UPDATED CODE SECTION: Changed to use detailed_df when available
        # Generate the Excel report - use detailed_df instead of filtered_df_display since it contains overdue data
        # If we have overdue analysis, use detailed_df which includes the overdue columns
        if 'detailed_df' in locals() and 'Overdue Jobs' in detailed_df.columns:
            export_df = detailed_df
        else:
            export_df = filtered_df_display
            
        excel_file = utils.create_excel_report(
            export_df, 
            analysis_results,
            file_level_overdue_data
        )
        
        # Create an appropriate filename
        report_filename = f"Job_Status_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        # Add overdue jobs info to the filename if available
        if analysis_results and 'overdue_jobs_count' in analysis_results and analysis_results['overdue_jobs_count'] > 0:
            report_filename = f"Job_Status_Report_with_Overdue_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        st.download_button(
            label="Download Excel Report",
            data=excel_file,
            file_name=report_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload one or more CSV files to analyze.")