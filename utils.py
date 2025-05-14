import pandas as pd
import os
import re
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import plotly.express as px
import plotly.graph_objects as go

def process_csv_file(file):
    """Process a single CSV file and extract relevant information."""
    # Initialize variables with default values
    filename = "Unknown"
    formatted_date = "Unknown"
    
    try:
        # Extract date from filename using regex
        filename = file.name
        date_match = re.search(r'\b(\d{2})(\d{2})(\d{4})\b', filename)
        formatted_date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}" if date_match else "Unknown"
        
        # Read CSV file
        df = pd.read_csv(file)
        
        # Identify the Vessel column
        vessel_column = next((col for col in df.columns if 'vessel' in col.lower()), None)
        vessel_name = df[vessel_column].iloc[0] if vessel_column else "Vessel column not found"
        
        # Identify the Job Status column
        status_column = next((col for col in df.columns if 'status' in col.lower()), None)
        
        # Count total jobs
        job_count = len(df)
        
        # Count jobs with Job Status == 'New'
        if status_column:
            df[status_column] = df[status_column].astype(str).str.strip()
            new_job_count = (df[status_column] == 'New').sum()
        else:
            new_job_count = 0
            
        return {
            'File Name': filename,
            'Vessel Name': vessel_name,
            'Total Count of Jobs': job_count,
            'New Job Count': new_job_count,
            'Date Extracted from File Name': formatted_date
        }
    except Exception as e:
        return {
            'File Name': filename,
            'Vessel Name': 'Error',
            'Total Count of Jobs': 'Error',
            'New Job Count': 'Error',
            'Date Extracted from File Name': formatted_date
        }

def create_vessel_job_distribution_chart(df, overdue_data=None):
    """Create a bar chart showing job distribution across vessels for individual files.
    
    Args:
        df: DataFrame with vessel job data
        overdue_data: Optional dictionary with overdue jobs data
    """
    # Sort data by date to maintain chronological order
    df = df.sort_values('Date Extracted from File Name')
    
    fig = go.Figure()
    
    # Add total jobs bars
    fig.add_trace(go.Bar(
        name='Total Jobs',
        x=[f"{row['Vessel Name']} - {row['File Name']}" for _, row in df.iterrows()],
        y=df['Total Count of Jobs'],
        marker_color='#1F4E78'
    ))
    
    # Add new jobs bars
    fig.add_trace(go.Bar(
        name='New Jobs',
        x=[f"{row['Vessel Name']} - {row['File Name']}" for _, row in df.iterrows()],
        y=df['New Job Count'],
        marker_color='#F63366'
    ))
    
    # Add overdue jobs bars if data is provided
    if overdue_data and 'file_results' in overdue_data and overdue_data['file_results']:
        # Create mappings of file names to overdue and critical overdue counts
        overdue_map = {}
        critical_map = {}
        
        # Process each file result from the analysis
        for file_result in overdue_data['file_results']:
            file_name = file_result['file_name']
            overdue_map[file_name] = file_result['overdue_jobs_count']
            critical_map[file_name] = file_result['critical_overdue_jobs_count']
        
        # Create lists for overdue jobs per vessel-file combination
        overdue_jobs = []
        critical_overdue_jobs = []
        
        # Match overdue data with the vessel-file combinations
        for _, row in df.iterrows():
            file_name = row['File Name']
            # Try to find an exact match
            if file_name in overdue_map:
                overdue_jobs.append(overdue_map[file_name])
                critical_overdue_jobs.append(critical_map[file_name])
            else:
                # Try basename matching or partial matching
                matched = False
                file_basename = os.path.basename(file_name)
                
                # Try matching with just the filename
                if file_basename in overdue_map:
                    overdue_jobs.append(overdue_map[file_basename])
                    critical_overdue_jobs.append(critical_map[file_basename])
                    matched = True
                else:
                    # Try partial matching
                    for analysis_file in overdue_map:
                        if file_name in analysis_file or analysis_file in file_name:
                            overdue_jobs.append(overdue_map[analysis_file])
                            critical_overdue_jobs.append(critical_map[analysis_file])
                            matched = True
                            break
                
                # If no match found, add zeros
                if not matched:
                    overdue_jobs.append(0)
                    critical_overdue_jobs.append(0)
        
        # Add overdue jobs bars for each vessel-file
        if any(overdue_jobs):
            fig.add_trace(go.Bar(
                name='Overdue Jobs',
                x=[f"{row['Vessel Name']} - {row['File Name']}" for _, row in df.iterrows()],
                y=overdue_jobs,
                marker_color='#FF9F40'
            ))
        
        # Add critical overdue jobs bars for each vessel-file
        if any(critical_overdue_jobs):
            fig.add_trace(go.Bar(
                name='Critical Overdue Jobs',
                x=[f"{row['Vessel Name']} - {row['File Name']}" for _, row in df.iterrows()],
                y=critical_overdue_jobs,
                marker_color='#FF5252'
            ))
    
    # Update layout with improved readability
    fig.update_layout(
        title='Job Distribution by Vessel and File',
        xaxis_title='Vessel - File',
        yaxis_title='Number of Jobs',
        barmode='group',
        height=500,  # Increased height for better visibility
        showlegend=True,
        xaxis=dict(
            tickangle=45,  # Angled labels for better readability
            tickmode='array',
            ticktext=[f"{row['Vessel Name']}<br>{row['File Name']}" for _, row in df.iterrows()],
            tickvals=list(range(len(df)))
        ),
        margin=dict(b=150)  # Increased bottom margin for rotated labels
    )
    
    return fig

def create_jobs_timeline_chart(df, overdue_data=None):
    """Create a line chart showing job trends over time.
    
    Args:
        df: DataFrame with job data
        overdue_data: Optional dictionary with overdue jobs data
    """
    timeline_data = df.groupby('Date Extracted from File Name').agg({
        'Total Count of Jobs': 'sum',
        'New Job Count': 'sum'
    }).reset_index()
    
    # Sort by date
    timeline_data['Date Extracted from File Name'] = pd.to_datetime(timeline_data['Date Extracted from File Name'])
    timeline_data = timeline_data.sort_values('Date Extracted from File Name')
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=timeline_data['Date Extracted from File Name'],
        y=timeline_data['Total Count of Jobs'],
        name='Total Jobs',
        line=dict(color='#1F4E78', width=2)
    ))
    fig.add_trace(go.Scatter(
        x=timeline_data['Date Extracted from File Name'],
        y=timeline_data['New Job Count'],
        name='New Jobs',
        line=dict(color='#F63366', width=2)
    ))
    
    # Add overdue jobs to the timeline if data exists
    if overdue_data and 'file_results' in overdue_data and overdue_data['file_results']:
        # Group data by date
        date_to_file = {}
        for _, row in df.iterrows():
            date = pd.to_datetime(row['Date Extracted from File Name'])
            file_name = row['File Name']
            if date not in date_to_file:
                date_to_file[date] = []
            date_to_file[date].append(file_name)
        
        # Create mappings for overdue data
        file_to_overdue = {}
        file_to_critical = {}
        for file_result in overdue_data['file_results']:
            file_to_overdue[file_result['file_name']] = file_result['overdue_jobs_count']
            file_to_critical[file_result['file_name']] = file_result['critical_overdue_jobs_count']
        
        # Create data for overdue and critical overdue by date
        dates = []
        overdue_by_date = []
        critical_by_date = []
        
        for date, files in date_to_file.items():
            date_overdue = 0
            date_critical = 0
            
            for file in files:
                # Try exact match
                if file in file_to_overdue:
                    date_overdue += file_to_overdue[file]
                    date_critical += file_to_critical[file]
                else:
                    # Try basename
                    basename = os.path.basename(file)
                    if basename in file_to_overdue:
                        date_overdue += file_to_overdue[basename]
                        date_critical += file_to_critical[basename]
                    else:
                        # Try partial matching
                        for analysis_file in file_to_overdue:
                            if file in analysis_file or analysis_file in file:
                                date_overdue += file_to_overdue[analysis_file]
                                date_critical += file_to_critical[analysis_file]
                                break
            
            dates.append(date)
            overdue_by_date.append(date_overdue)
            critical_by_date.append(date_critical)
        
        # Sort the data by date
        sorted_indices = [i for i, _ in sorted(enumerate(dates), key=lambda x: x[1])]
        sorted_dates = [dates[i] for i in sorted_indices]
        sorted_overdue = [overdue_by_date[i] for i in sorted_indices]
        sorted_critical = [critical_by_date[i] for i in sorted_indices]
        
        # Add overdue jobs line
        if any(sorted_overdue):
            fig.add_trace(go.Scatter(
                x=sorted_dates,
                y=sorted_overdue,
                name='Overdue Jobs',
                line=dict(color='#FF9F40', width=2, dash='dot'),
                mode='lines+markers+text',
                text=sorted_overdue,
                textposition="top center"
            ))
        
        # Add critical overdue jobs line
        if any(sorted_critical):
            fig.add_trace(go.Scatter(
                x=sorted_dates,
                y=sorted_critical,
                name='Critical Overdue',
                line=dict(color='#FF5252', width=2, dash='dot'),
                mode='lines+markers+text',
                text=sorted_critical,
                textposition="top right"
            ))
    
    fig.update_layout(
        title='Job Trends Over Time',
        xaxis_title='Date',
        yaxis_title='Number of Jobs',
        height=400,
        showlegend=True
    )
    return fig

def create_jobs_pie_chart(df, overdue_data=None):
    """Create a pie chart showing the proportion of job statuses.
    
    Args:
        df: DataFrame with job data
        overdue_data: Optional dictionary with overdue jobs data
    """
    # Calculate base metrics
    total_jobs = df['Total Count of Jobs'].sum()
    new_jobs = df['New Job Count'].sum()
    
    # Calculate overdue and critical overdue values from file-level data
    overdue_jobs = 0
    critical_overdue = 0
    
    if overdue_data and 'file_results' in overdue_data and overdue_data['file_results']:
        # Create mappings for overdue data
        file_to_overdue = {}
        file_to_critical = {}
        for file_result in overdue_data['file_results']:
            file_to_overdue[file_result['file_name']] = file_result['overdue_jobs_count']
            file_to_critical[file_result['file_name']] = file_result['critical_overdue_jobs_count']
        
        # Match and sum overdue data for files in the current DataFrame
        for _, row in df.iterrows():
            file_name = row['File Name']
            # Try exact match
            if file_name in file_to_overdue:
                overdue_jobs += file_to_overdue[file_name]
                critical_overdue += file_to_critical[file_name]
            else:
                # Try basename
                basename = os.path.basename(file_name)
                if basename in file_to_overdue:
                    overdue_jobs += file_to_overdue[basename]
                    critical_overdue += file_to_critical[basename]
                else:
                    # Try partial matching
                    for analysis_file in file_to_overdue:
                        if file_name in analysis_file or analysis_file in file_name:
                            overdue_jobs += file_to_overdue[analysis_file]
                            critical_overdue += file_to_critical[analysis_file]
                            break
    
    # If we have overdue jobs data, include it in the chart
    if overdue_jobs > 0:
        # Calculate the remaining jobs (ensuring it's not negative)
        remaining_jobs = max(0, total_jobs - new_jobs - overdue_jobs)
        
        # If we have critical overdue jobs, show them separately
        if critical_overdue > 0:
            labels = ['New Jobs', 'Overdue Jobs', 'Critical Overdue', 'Other Jobs']
            values = [new_jobs, overdue_jobs - critical_overdue, critical_overdue, remaining_jobs]
            colors = ['#F63366', '#FF9F40', '#FF5252', '#1F4E78']
        else:
            labels = ['New Jobs', 'Overdue Jobs', 'Other Jobs']
            values = [new_jobs, overdue_jobs, remaining_jobs]
            colors = ['#F63366', '#FF9F40', '#1F4E78']
        
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=.4,
            marker_colors=colors
        )])
        
        fig.update_layout(
            title='Job Status Distribution',
            height=400,
            showlegend=True
        )
    else:
        # Original pie chart without overdue data
        existing_jobs = total_jobs - new_jobs
        
        fig = go.Figure(data=[go.Pie(
            labels=['New Jobs', 'Existing Jobs'],
            values=[new_jobs, existing_jobs],
            hole=.4,
            marker_colors=['#F63366', '#1F4E78']
        )])
        
        fig.update_layout(
            title='New vs. Existing Jobs Distribution',
            height=400,
            showlegend=True
        )
    
    return fig

def create_overdue_jobs_chart(df, overdue_data=None):
    """Create a bar chart showing overdue and critical overdue jobs comparison.
    
    Args:
        df: DataFrame with job data
        overdue_data: Optional dictionary with overdue jobs data
    """
    # Calculate total overdue and critical overdue jobs from file-level data
    total_overdue = 0
    total_critical = 0
    
    if overdue_data and 'file_results' in overdue_data and overdue_data['file_results']:
        # Sum up all overdue and critical overdue jobs across all files
        for file_result in overdue_data['file_results']:
            total_overdue += file_result['overdue_jobs_count']
            total_critical += file_result['critical_overdue_jobs_count']
    
    # Create the bar chart
    fig = go.Figure()
    
    # Add bars for overdue and critical overdue
    fig.add_trace(go.Bar(
        x=['Overdue Jobs', 'Critical Overdue Jobs'],
        y=[total_overdue, total_critical],
        marker_color=['#FF9F40', '#FF5252']
    ))
    
    # Update layout for better readability
    fig.update_layout(
        title='Overdue Jobs Analysis',
        yaxis_title='Count',
        xaxis_title='Job Type',
        height=400,
        showlegend=False
    )
    
    return fig

def analyze_overdue_jobs(df):
    """Analyze overdue jobs and critical overdue jobs from a DataFrame.
    
    Returns a dictionary with overdue job metrics per individual file/record.
    """
    try:
        # Make a copy to avoid modifying the original DataFrame
        df_copy = df.copy()
        
        # Clean column names
        df_copy.columns = df_copy.columns.str.strip()
        
        # Initialize results for each file
        file_results = []
        
        # Process each row as an individual file record
        if 'Calculated Due Date' in df_copy.columns and 'Job Status' in df_copy.columns:
            # Convert 'Calculated Due Date' to datetime once
            df_copy['Calculated Due Date'] = pd.to_datetime(df_copy['Calculated Due Date'], errors='coerce')
            
            # Define today's date
            today = pd.to_datetime(datetime.today().date())
            
            # We'll create a new DataFrame to store individual file results
            result_df = pd.DataFrame()
            
            # Check for the source file column first (added during processing in app.py)
            if '_source_file' in df_copy.columns:
                files = df_copy['_source_file'].unique()
                file_col = '_source_file'
            # Otherwise check for a File Name column
            elif 'File Name' in df_copy.columns:
                files = df_copy['File Name'].unique()
                file_col = 'File Name'
            else:
                # Analyze as one file if no file name column
                files = ['Entire Dataset']
                df_copy['_file_id'] = 'Entire Dataset'
                file_col = '_file_id'
            
            # Process each file
            for file_name in files:
                file_data = df_copy[df_copy[file_col] == file_name]
                
                # Calculate overdue jobs for this file
                overdue_jobs = file_data[
                    (file_data['Calculated Due Date'] <= today) &
                    (file_data['Job Status'].str.strip().str.lower().isin(['pending', 'in progress on board']))
                ]
                overdue_jobs_count = len(overdue_jobs)
                
                # Calculate critical overdue jobs for this file
                try:
                    # Check if the column exists first
                    if 'Unnamed: 0' in file_data.columns:
                        critical_overdue_jobs = file_data[
                            (file_data['Unnamed: 0'].astype(str).str.strip().str.lower() == 'c') &
                            (file_data['Calculated Due Date'] <= today) &
                            (file_data['Job Status'].str.strip().str.lower().isin(['pending', 'in progress on board']))
                        ]
                    else:
                        # Try alternative columns that might indicate criticality
                        critical_col = next((col for col in file_data.columns if 'critical' in col.lower() or 'priority' in col.lower()), None)
                        
                        if critical_col:
                            critical_overdue_jobs = file_data[
                                (file_data[critical_col].astype(str).str.strip().str.lower().isin(['c', 'critical', 'high', 'yes', 'true'])) &
                                (file_data['Calculated Due Date'] <= today) &
                                (file_data['Job Status'].str.strip().str.lower().isin(['pending', 'in progress on board']))
                            ]
                        else:
                            # No criticality column found
                            critical_overdue_jobs = pd.DataFrame()
                except Exception as e:
                    print(f"Error processing critical jobs for {file_name}: {str(e)}")
                    critical_overdue_jobs = pd.DataFrame()
                
                critical_overdue_jobs_count = len(critical_overdue_jobs)
                
                # Total jobs in this file
                total_jobs = len(file_data)
                
                # Calculate percentages for this file
                if total_jobs > 0:
                    overdue_jobs_percentage = round((overdue_jobs_count / total_jobs) * 100, 2)
                    critical_overdue_jobs_percentage = round((critical_overdue_jobs_count / total_jobs) * 100, 2)
                else:
                    overdue_jobs_percentage = 0
                    critical_overdue_jobs_percentage = 0
                
                # Add to results
                file_results.append({
                    'file_name': file_name,
                    'total_jobs': total_jobs,
                    'overdue_jobs_count': overdue_jobs_count,
                    'overdue_jobs_percentage': overdue_jobs_percentage,
                    'critical_overdue_jobs_count': critical_overdue_jobs_count,
                    'critical_overdue_jobs_percentage': critical_overdue_jobs_percentage,
                    'overdue_jobs': overdue_jobs,
                    'critical_overdue_jobs': critical_overdue_jobs
                })
            
            # Create DataFrame from results
            results_df = pd.DataFrame(file_results)
            
            # Calculate overall totals
            total_all_jobs = results_df['total_jobs'].sum()
            total_overdue = results_df['overdue_jobs_count'].sum()
            total_critical = results_df['critical_overdue_jobs_count'].sum()
            
            if total_all_jobs > 0:
                overall_overdue_pct = round((total_overdue / total_all_jobs) * 100, 2)
                overall_critical_pct = round((total_critical / total_all_jobs) * 100, 2)
            else:
                overall_overdue_pct = 0
                overall_critical_pct = 0
            
            # Combine all overdue and critical jobs
            all_overdue = pd.concat([result['overdue_jobs'] for result in file_results]) if file_results else pd.DataFrame()
            all_critical = pd.concat([result['critical_overdue_jobs'] for result in file_results]) if file_results else pd.DataFrame()
            
            return {
                'file_results': file_results,  # Individual file results
                'overdue_jobs_count': total_overdue,  # Total across all files
                'overdue_jobs_percentage': overall_overdue_pct,
                'critical_overdue_jobs_count': total_critical,
                'critical_overdue_jobs_percentage': overall_critical_pct,
                'total_jobs': total_all_jobs,
                'overdue_jobs': all_overdue,  # All overdue jobs combined
                'critical_overdue_jobs': all_critical  # All critical jobs combined
            }
        else:
            return {
                'file_results': [],
                'overdue_jobs_count': 0,
                'overdue_jobs_percentage': 0,
                'critical_overdue_jobs_count': 0,
                'critical_overdue_jobs_percentage': 0,
                'total_jobs': 0,
                'overdue_jobs': pd.DataFrame(),
                'critical_overdue_jobs': pd.DataFrame()
            }
    except Exception as e:
        print(f"Error analyzing overdue jobs: {str(e)}")
        return {
            'file_results': [],
            'overdue_jobs_count': 0,
            'overdue_jobs_percentage': 0,
            'critical_overdue_jobs_count': 0,
            'critical_overdue_jobs_percentage': 0,
            'total_jobs': 0,
            'overdue_jobs': pd.DataFrame(),
            'critical_overdue_jobs': pd.DataFrame()
        }

def create_overdue_jobs_chart(overdue_data, critical_data):
    """Create a bar chart comparing overdue and critical overdue jobs."""
    labels = ['Overdue Jobs', 'Critical Overdue Jobs']
    values = [overdue_data, critical_data]
    
    fig = go.Figure(data=[
        go.Bar(name='Count', x=labels, y=values, 
               marker_color=['#FF9F40', '#FF5252'])
    ])
    
    fig.update_layout(
        title='Overdue Jobs Analysis',
        xaxis_title='Job Type',
        yaxis_title='Count',
        height=400,
        showlegend=False
    )
    
    return fig

def create_excel_report(df, overdue_data=None, file_level_overdue_data=None):
    """Create a formatted Excel report from the DataFrame.
    
    Args:
        df: DataFrame with job data
        overdue_data: Optional dictionary with overdue jobs data
        file_level_overdue_data: Optional DataFrame with file-level overdue metrics
    """
    output = BytesIO()
    
    # Define the exact column order as shown in the image
    required_columns = [
        'File Name',
        'Vessel Name',
        'Total Count of Jobs',
        'New Job Count',
        'Date Extracted from File Name',
        'Overdue Jobs',
        'Critical Overdue',
        'Overdue %'
    ]
    
    # Filter the DataFrame to include only the required columns that exist
    available_columns = [col for col in required_columns if col in df.columns]
    df = df[available_columns]
    
    # Save DataFrame to Excel
    df.to_excel(output, index=False)
    
    # Load workbook for formatting
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active
    
    # Rename the main worksheet
    ws.title = "Job Status Summary"
    
    # Define styles
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Cell borders
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Format headers
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Format data cells
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
    
    # Define orange fill for conditional formatting (duplicates)
    orange_fill = PatternFill(start_color="FFB266", end_color="FFB266", fill_type="solid")
    dxf = DifferentialStyle(fill=orange_fill)
    
    # Create rule for duplicate values in Vessel Name column
    dup_rule = Rule(type="duplicateValues", dxf=dxf, stopIfTrue=False)
    ws.conditional_formatting.add(f'B2:B{ws.max_row}', dup_rule)
    
    # Alternating row colors
    gray_fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
    for row in range(2, ws.max_row + 1, 2):
        for cell in ws[row]:
            cell.fill = gray_fill
    
    # Add Excel table with formatting - dynamically determine number of columns
    # Get the maximum column index (letter)
    max_col = ws.max_column
    max_col_letter = get_column_letter(max_col)
    table_ref = f"A1:{max_col_letter}{ws.max_row}"
    table = Table(displayName="JobSummaryTable", ref=table_ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    ws.add_table(table)
    
    # Adjust column widths
    for column in ws.columns:
        max_length = 0
        column = list(column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width
    
    # No additional sheets are being created as per requirements
    
    # Save to BytesIO
    output_formatted = BytesIO()
    wb.save(output_formatted)
    output_formatted.seek(0)
    
    return output_formatted
