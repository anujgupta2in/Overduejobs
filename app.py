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
    - Overdue job analysis
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
                'Overdue %': f"{file_result['overdue_jobs_percentage']}%",
                'Critical %': f"{file_result['critical_overdue_jobs_percentage']}%"
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
        critical_pct = []
        
        for _, row in job_status_table.iterrows():
            file_name = row['File Name']
            simple_name = os.path.basename(file_name)
            
            # Try exact match first
            if file_name in simplified_map:
                overdue_jobs.append(simplified_map[file_name]['Overdue Jobs'])
                critical_overdue.append(simplified_map[file_name]['Critical Overdue'])
                overdue_pct.append(simplified_map[file_name]['Overdue %'])
                critical_pct.append(simplified_map[file_name]['Critical %'])
            # Try with just the filename
            elif simple_name in simplified_map:
                overdue_jobs.append(simplified_map[simple_name]['Overdue Jobs'])
                critical_overdue.append(simplified_map[simple_name]['Critical Overdue'])
                overdue_pct.append(simplified_map[simple_name]['Overdue %'])
                critical_pct.append(simplified_map[simple_name]['Critical %'])
            else:
                # Try partial matching as a last resort
                matched = False
                for analysis_file in file_analysis_map:
                    # Check if analysis file contains this file name or vice versa
                    if file_name in analysis_file or analysis_file in file_name:
                        overdue_jobs.append(file_analysis_map[analysis_file]['Overdue Jobs'])
                        critical_overdue.append(file_analysis_map[analysis_file]['Critical Overdue'])
                        overdue_pct.append(file_analysis_map[analysis_file]['Overdue %'])
                        critical_pct.append(file_analysis_map[analysis_file]['Critical %'])
                        matched = True
                        break
                
                if not matched:
                    # Nothing matched, put placeholder values
                    overdue_jobs.append("N/A")
                    critical_overdue.append("N/A")
                    overdue_pct.append("N/A")
                    critical_pct.append("N/A")
        
        # Add the overdue metrics to the table
        job_status_table['Overdue Jobs'] = overdue_jobs
        job_status_table['Critical Overdue'] = critical_overdue
        job_status_table['Overdue %'] = overdue_pct
        job_status_table['Critical %'] = critical_pct
    else:
        # Add placeholders if no overdue analysis is available
        job_status_table['Overdue Jobs'] = "N/A"
        job_status_table['Critical Overdue'] = "N/A"
        job_status_table['Overdue %'] = "N/A"
        job_status_table['Critical %'] = "N/A"
    
    # Define a function to color cells based on percentage values
    def highlight_percentage(val):
        if isinstance(val, str) and val != "N/A":
            try:
                # Remove the % sign and convert to float
                num_val = float(val.replace('%', ''))
                if num_val > 3.0:
                    return 'background-color: #FF4B4B'  # Red background for > 3%
            except ValueError:
                pass
        return ''
    
    # Apply the styling to the dataframe for both percentage columns
    styled_job_status_table = job_status_table.style.applymap(
        highlight_percentage, 
        subset=['Critical %']
    ).applymap(
        highlight_percentage,
        subset=['Overdue %']
    )
    
    # Display the table with styling
    st.dataframe(
        styled_job_status_table,
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
        
        # Total Jobs - Blue color
        overdue_metrics[0].metric(
            "Total Jobs", 
            analysis_results['total_jobs'],
            delta=None,
            delta_color="normal",
            help=None,
            label_visibility="visible"
        )
        # Apply blue color with HTML
        overdue_metrics[0].markdown(
            "<style>div[data-testid='stMetric']:nth-child(1) > div:nth-child(1) > p { color: #1E88E5; font-weight: bold; }</style>", 
            unsafe_allow_html=True
        )
        
        # Calculate new jobs in the detailed file (if available)
        new_status_jobs = 0
        if 'overdue_jobs' in analysis_results and not analysis_results['overdue_jobs'].empty:
            if 'Job Status' in analysis_results['overdue_jobs'].columns:
                new_status_jobs = (analysis_results['overdue_jobs']['Job Status'].str.strip().str.lower() == 'new').sum()
        
        # New Jobs - Green color
        overdue_metrics[1].metric(
            "New Status Jobs", 
            new_status_jobs
        )
        # Apply green color with HTML
        overdue_metrics[1].markdown(
            "<style>div[data-testid='stMetric']:nth-child(2) > div:nth-child(1) > p { color: #4CAF50; font-weight: bold; }</style>", 
            unsafe_allow_html=True
        )
        
        # Overdue Jobs - Orange color
        overdue_metrics[2].metric(
            "Total Overdue", 
            analysis_results['overdue_jobs_count']
        )
        # Apply orange color with HTML
        overdue_metrics[2].markdown(
            "<style>div[data-testid='stMetric']:nth-child(3) > div:nth-child(1) > p { color: #FF9800; font-weight: bold; }</style>", 
            unsafe_allow_html=True
        )
        
        # Critical Overdue - Red color
        overdue_metrics[3].metric(
            "Critical Overdue", 
            analysis_results['critical_overdue_jobs_count']
        )
        # Apply red color with HTML
        overdue_metrics[3].markdown(
            "<style>div[data-testid='stMetric']:nth-child(4) > div:nth-child(1) > p { color: #F44336; font-weight: bold; }</style>", 
            unsafe_allow_html=True
        )
        
        # Overdue Percentage - Orange color
        overdue_metrics[4].metric(
            "Overdue %", 
            f"{analysis_results['overdue_jobs_percentage']:.1f}%"
        )
        # Apply orange color with HTML
        overdue_metrics[4].markdown(
            "<style>div[data-testid='stMetric']:nth-child(5) > div:nth-child(1) > p { color: #FF9800; font-weight: bold; }</style>", 
            unsafe_allow_html=True
        )

    # Clear status text after processing
    status_text.empty()
    progress_bar.empty()

    # Display summary table
    st.subheader("📋 File Summary")
    st.dataframe(
        filtered_df[['File Name', 'Vessel Name', 'Total Count of Jobs', 'New Job Count', 'Date Extracted from File Name']],
        use_container_width=True,
        hide_index=True
    )

    # Visualization Section
    st.subheader("📊 Data Visualizations")

    # Create tabs for different visualizations
    tab1, tab2, tab3 = st.tabs(["📊 Job Distribution", "📈 Timeline Trends", "🥧 Job Status Pie Chart"])

    with tab1:
        if len(filtered_df) > 0:
            fig_bar = utils.create_vessel_job_distribution_chart(filtered_df, analysis_results)
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("No data available for the selected filters.")

    with tab2:
        if len(filtered_df) > 0:
            fig_line = utils.create_jobs_timeline_chart(filtered_df, analysis_results)
            st.plotly_chart(fig_line, use_container_width=True)
        else:
            st.info("No data available for the selected filters.")

    with tab3:
        if len(filtered_df) > 0:
            fig_pie = utils.create_jobs_pie_chart(filtered_df, analysis_results)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No data available for the selected filters.")

    # Show detailed overdue jobs if available
    if analysis_results and 'overdue_jobs' in analysis_results and not analysis_results['overdue_jobs'].empty:
        st.subheader("⚠️ Detailed Overdue Jobs")
        
        # Add expandable sections for different overdue categories
        with st.expander("View All Overdue Jobs", expanded=False):
            st.dataframe(
                analysis_results['overdue_jobs'],
                use_container_width=True,
                hide_index=True
            )
        
        if 'critical_overdue_jobs' in analysis_results and not analysis_results['critical_overdue_jobs'].empty:
            with st.expander("View Critical Overdue Jobs", expanded=False):
                st.dataframe(
                    analysis_results['critical_overdue_jobs'],
                    use_container_width=True,
                    hide_index=True
                )

    # Excel Export Section
    st.subheader("📤 Export Data")
    
    if st.button("Generate Excel Report", type="primary"):
        try:
            with st.spinner("Generating Excel report..."):
                # Generate Excel report
                excel_buffer = utils.create_excel_report(filtered_df, analysis_results)
                
                # Create download button
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_buffer.getvalue(),
                    file_name=f"job_status_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Excel report generated successfully!")
        except Exception as e:
            st.error(f"Error generating Excel report: {str(e)}")

else:
    # Show instructions when no files are uploaded
    st.info("""
        👆 Upload one or more CSV files to get started.
        
        **Expected CSV file format:**
        - Must contain job status information
        - Should include vessel names
        - Filenames should contain dates in format DDMMYYYY
        - For overdue analysis, include 'Calculated Due Date' and 'Job Status' columns
    """)
