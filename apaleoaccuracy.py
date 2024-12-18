import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.graph_objs as go
from plotly.subplots import make_subplots
from io import BytesIO

# Set page layout to wide
st.set_page_config(layout="wide", page_title="Guestline Daily Variance and Accuracy Calculator")

# Define the function to read the two CSV files
def read_data(file, is_history_and_forecast_file):
    if not file.name.endswith('.csv'):
        raise ValueError("Unsupported file format. Please upload a .csv file.")

    if is_history_and_forecast_file:
        # History and Forecast file
        expected_columns = ['businessDay', 'soldCount', 'noShowsCount', 'netAccommodationRevenue', 'netRevenue', 'grossRevenue']
        df = pd.read_csv(file)
        for col in expected_columns:
            if col not in df.columns:
                raise ValueError(f"Expected column '{col}' not found in the uploaded file.")
        df = df[['businessDay', 'soldCount', 'netAccommodationRevenue']]
        df.columns = ['date', 'AF RNs', 'AF Rev']  # Rename columns for consistency
        try:
            df['date'] = pd.to_datetime(df['date'], format='%Y-%m-%d', errors='coerce').dt.date
        except Exception as e:
            raise ValueError(f"Error converting 'businessDay' column to datetime: {e}")
    else:
        # Daily totals file
        expected_columns = ['arrivalDate', '"rn"', '"revNet"', '"revTotal"', '"revFb"', '"revResTotal"']
        df = pd.read_csv(file, delimiter=';', quotechar='"')
        for col in expected_columns:
            if col not in df.columns:
                raise ValueError(f"Expected column '{col}' not found in the uploaded file.")
        df = df[['arrivalDate', '"rn"', '"revNet"']]
        df.columns = ['date', 'Juyo RN', 'Juyo Rev']  # Rename columns for consistency
        try:
            df['date'] = pd.to_datetime(df['date'], format='%Y-%m-%d', errors='coerce').dt.date
        except Exception as e:
            raise ValueError(f"Error converting 'arrivalDate' column to datetime: {e}")

    if df['date'].isnull().any():
        raise ValueError("Some dates could not be parsed. Please ensure that the date column is in a valid date format.")

    return df


# Define color coding for accuracy values
def color_scale(val):
    """Color scale for percentages."""
    if isinstance(val, str) and '%' in val:
        val = float(val.strip('%'))
        if val >= 98:
            return 'background-color: #469798; color: white;'  # green
        elif 95 <= val < 98:
            return 'background-color: #F2A541; color: white;'  # yellow
        else:
            return 'background-color: #BF3100; color: white;'  # red
    return ''
    
# Function to create Excel file for download with color formatting and accuracy matrix
def create_excel_download(combined_df, base_filename, past_accuracy_rn, past_accuracy_rev, future_accuracy_rn, future_accuracy_rev):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Write the Accuracy Matrix
        accuracy_matrix = pd.DataFrame({
            'Metric': ['RNs', 'Revenue'],
            'Past': [past_accuracy_rn / 100, past_accuracy_rev / 100],  # Store as decimal
            'Future': [future_accuracy_rn / 100, future_accuracy_rev / 100]  # Store as decimal
        })
        
        accuracy_matrix.to_excel(writer, sheet_name='Accuracy Matrix', index=False, startrow=1)
        worksheet_accuracy = writer.sheets['Accuracy Matrix']

        # Define custom formats and colors
        format_green = workbook.add_format({'bg_color': '#469798', 'font_color': '#FFFFFF'})
        format_yellow = workbook.add_format({'bg_color': '#F2A541', 'font_color': '#FFFFFF'})
        format_red = workbook.add_format({'bg_color': '#BF3100', 'font_color': '#FFFFFF'})
        format_percent = workbook.add_format({'num_format': '0.00%'})  # Percentage format

        # Apply percentage format to the relevant cells in Accuracy Matrix
        worksheet_accuracy.set_column('B:C', None, format_percent)  # Set percentage format for both Past and Future columns

        # Apply conditional formatting for Accuracy Matrix
        worksheet_accuracy.conditional_format('B3:B4', {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
        worksheet_accuracy.conditional_format('B3:B4', {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
        worksheet_accuracy.conditional_format('B3:B4', {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})
        worksheet_accuracy.conditional_format('C3:C4', {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
        worksheet_accuracy.conditional_format('C3:C4', {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
        worksheet_accuracy.conditional_format('C3:C4', {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

        # Write the combined past and future results to a single sheet
        if not combined_df.empty:
            # Ensure percentage columns are properly formatted as decimals
            combined_df['Abs RN Accuracy'] = combined_df['Abs RN Accuracy'].str.rstrip('%').astype('float') / 100
            combined_df['Abs Rev Accuracy'] = combined_df['Abs Rev Accuracy'].str.rstrip('%').astype('float') / 100

            combined_df.to_excel(writer, sheet_name='Daily Variance Detail', index=False)
            worksheet_combined = writer.sheets['Daily Variance Detail']

            # Define number formats
            format_number = workbook.add_format({'num_format': '#,##0.00'})  # Floats
            format_whole = workbook.add_format({'num_format': '0'})  # Whole numbers

            # Format columns in the "Daily Variance Detail" sheet
            worksheet_combined.set_column('A:A', None, format_whole)  # Date
            worksheet_combined.set_column('B:B', None, format_whole)  # AF RNs
            worksheet_combined.set_column('C:C', None, format_number)  # AF Rev
            worksheet_combined.set_column('D:D', None, format_whole)  # Juyo RN
            worksheet_combined.set_column('E:E', None, format_number)  # Juyo Rev
            worksheet_combined.set_column('F:F', None, format_whole)  # RN Diff
            worksheet_combined.set_column('G:G', None, format_number)  # Rev Diff
            worksheet_combined.set_column('H:H', None, format_percent)  # Abs RN Accuracy
            worksheet_combined.set_column('I:I', None, format_percent)  # Abs Rev Accuracy

            # Apply conditional formatting to the percentage columns (H and I)
            worksheet_combined.conditional_format('H2:H{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
            worksheet_combined.conditional_format('H2:H{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
            worksheet_combined.conditional_format('H2:H{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})
            worksheet_combined.conditional_format('I2:I{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
            worksheet_combined.conditional_format('I2:I{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
            worksheet_combined.conditional_format('I2:I{}'.format(len(combined_df) + 1),
                                                  {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

    output.seek(0)
    return output, base_filename

# Streamlit application
def main():
    # Center the title using markdown with HTML
    st.markdown("<h1 style='text-align: center;'> Guestline Daily Variance and Accuracy Calculator</h1>", unsafe_allow_html=True)
    # Note/warning box with matching colors
    st.warning("The reference date of the Daily Totals Extract should be equal to the latest Guestline file date.")
    
    # File uploaders in columns
    col1, col2 = st.columns(2)
    with col1:
        data_file = st.file_uploader("Upload History and Forecast (.xlsx or .csv)", type=['xlsx', 'csv'])
    with col2:
        csv_file = st.file_uploader("Upload Daily Totals Extract from Support UI", type=['csv'])

    # Only proceed with further checks if an xlsx file is uploaded
    if data_file and data_file.name.endswith('.xlsx'):
        try:
            # Check the number of sheets in the uploaded xlsx file
            excel_file = pd.ExcelFile(data_file)
            sheet_count = len(excel_file.sheet_names)

            # Display hotel code input and success message only if the xlsx file contains multiple sheets
            if sheet_count > 1:
                hotel_code = st.text_input("Enter Hotel Code (Sheet Name) - optional if the file contains a single sheet:")
                st.write("Files uploaded successfully. You can now choose to define the Hotel Code or proceed with processing the data.")
        except Exception as e:
            st.error(f"Error processing the xlsx file: {e}")
            return

    # When both files are uploaded
    if data_file and csv_file:
        # Display the process button
        if st.button("Process Data"):
            # Initialize progress bar
            progress_bar = st.progress(0)
            
            try:
                # Read the data from either xlsx or csv
                xlsx_df = read_data(data_file, hotel_code if 'hotel_code' in locals() else None)
            except ValueError as e:
                st.error(f"Error reading data file: {e}")
                return
            
            # Update progress bar
            progress_bar.progress(25)
            
            # Process CSV
            csv_df = pd.read_csv(csv_file, delimiter=';', quotechar='"')

            # Drop any columns that are completely empty (all NaN)
            csv_df = csv_df.dropna(axis=1, how='all')

            csv_df.columns = [col.replace('"', '').strip() for col in csv_df.columns]
            csv_df['arrivalDate'] = pd.to_datetime(csv_df['arrivalDate'], errors='coerce')
            csv_df['Juyo RN'] = csv_df['rn'].astype(int)
            csv_df['Juyo Rev'] = csv_df['revNet'].astype(float)
            
            # Ensure both date columns are of type datetime64[ns] before merging
            xlsx_df['date'] = pd.to_datetime(xlsx_df['date'], errors='coerce')
            csv_df['arrivalDate'] = pd.to_datetime(csv_df['arrivalDate'], errors='coerce')

            # Update progress bar
            progress_bar.progress(50)

            # Merge data
            merged_df = pd.merge(xlsx_df, csv_df, left_on='date', right_on='arrivalDate')
            
            # Calculate discrepancies for rooms and revenue
            merged_df['RN Diff'] = merged_df['Juyo RN'] - merged_df['AF RNs']
            merged_df['Rev Diff'] = merged_df['Juyo Rev'] - merged_df['AF Rev']
            
            # Calculate absolute accuracy percentages with handling for 0/0 cases
            merged_df['Abs RN Accuracy'] = merged_df.apply(
                lambda row: 100.0 if row['AF RNs'] == 0 and row['Juyo RN'] == 0 else 
                            (1 - abs(row['RN Diff']) / row['AF RNs']) * 100 if row['AF RNs'] != 0 else 0.0,
                axis=1
            )
            
            merged_df['Abs Rev Accuracy'] = merged_df.apply(
                lambda row: 100.0 if row['AF Rev'] == 0 and row['Juyo Rev'] == 0 else 
                            (1 - abs(row['Rev Diff']) / row['AF Rev']) * 100 if row['AF Rev'] != 0 else 0.0,
                axis=1
            )

            # Update progress bar
            progress_bar.progress(75)
            
            # Format accuracy percentages as strings with '%' symbol
            merged_df['Abs RN Accuracy'] = merged_df['Abs RN Accuracy'].map(lambda x: f"{x:.2f}%")
            merged_df['Abs Rev Accuracy'] = merged_df['Abs Rev Accuracy'].map(lambda x: f"{x:.2f}%")
            
            # Calculate overall accuracies
            current_date = pd.to_datetime('today').normalize()  # Get the current date without the time part
            past_mask = merged_df['date'] < current_date
            future_mask = merged_df['date'] >= current_date
            past_rooms_accuracy = (1 - (abs(merged_df.loc[past_mask, 'RN Diff']).sum() / merged_df.loc[past_mask, 'AF RNs'].sum())) * 100
            past_revenue_accuracy = (1 - (abs(merged_df.loc[past_mask, 'Rev Diff']).sum() / merged_df.loc[past_mask, 'AF Rev'].sum())) * 100
            future_rooms_accuracy = (1 - (abs(merged_df.loc[future_mask, 'RN Diff']).sum() / merged_df.loc[future_mask, 'AF RNs'].sum())) * 100
            future_revenue_accuracy = (1 - (abs(merged_df.loc[future_mask, 'Rev Diff']).sum() / merged_df.loc[future_mask, 'AF Rev'].sum())) * 100
            
            # Display accuracy matrix in a table within a container for width control
            accuracy_data = {
                "RNs": [f"{past_rooms_accuracy:.2f}%", f"{future_rooms_accuracy:.2f}%"],
                "Revenue": [f"{past_revenue_accuracy:.2f}%", f"{future_revenue_accuracy:.2f}%"]
            }
            accuracy_df = pd.DataFrame(accuracy_data, index=["Past", "Future"])
            
            # Center the accuracy matrix table
            with st.container():
                st.table(accuracy_df.style.applymap(color_scale).set_table_styles([{"selector": "th", "props": [("backgroundColor", "#f0f2f6")]}]))
            
            # Warning about future discrepancies with matching colors
            st.warning("Future discrepancies might be a result of timing discrepancies between the moment that the data was received and the moment that the history and forecast file was received.")
            
            # Interactive bar and line chart for visualizing discrepancies
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            
            # RN Discrepancies - Bar chart
            fig.add_trace(go.Bar(
                x=merged_df['date'],
                y=merged_df['RN Diff'],
                name='RNs Discrepancy',
                marker_color='#469798'
            ), secondary_y=False)
            
            # Revenue Discrepancies - Line chart
            fig.add_trace(go.Scatter(
                x=merged_df['date'],
                y=merged_df['Rev Diff'],
                name='Revenue Discrepancy',
                mode='lines+markers',
                line=dict(color='#BF3100', width=2),
                marker=dict(size=8)
            ), secondary_y=True)
            
            # Update plot layout for dynamic axis scaling and increased height
            max_room_discrepancy = merged_df['RN Diff'].abs().max()
            max_revenue_discrepancy = merged_df['Rev Diff'].abs().max()
            fig.update_layout(
                height=600,
                title='RNs and Revenue Discrepancy Over Time',
                xaxis_title='Date',
                yaxis_title='RNs Discrepancy',
                yaxis2_title='Revenue Discrepancy',
                yaxis=dict(range=[-max_room_discrepancy, max_room_discrepancy]),
                yaxis2=dict(range=[-max_revenue_discrepancy, max_revenue_discrepancy]),
                legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)
            )
            
            # Align grid lines
            fig.update_yaxes(matches=None, showgrid=True, gridwidth=1, gridcolor='grey')
            st.plotly_chart(fig, use_container_width=True)
            
            # Display daily variance detail in a table
            st.markdown("### Daily Variance Detail", unsafe_allow_html=True)
            detail_container = st.container()
            with detail_container:
                formatted_df = merged_df[['date', 'AF RNs', 'AF Rev', 'Juyo RN', 'Juyo Rev', 'RN Diff', 'Rev Diff', 'Abs RN Accuracy', 'Abs Rev Accuracy']]
                styled_df = formatted_df.style.applymap(color_scale, subset=['Abs RN Accuracy', 'Abs Rev Accuracy']).set_properties(**{'text-align': 'center'})
                st.table(styled_df)
            
            # Combine past and future data for export
            combined_df = pd.concat([formatted_df[past_mask], formatted_df[future_mask]])
            
            # Extract the filename prefix from the uploaded CSV
            csv_filename = csv_file.name.split('_')[0]
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            base_filename = f"{csv_filename}_AccuracyCheck_{current_time}"
            
            # Add Excel export functionality
            output, filename = create_excel_download(combined_df, base_filename,
                                                     past_rooms_accuracy, past_revenue_accuracy,
                                                     future_rooms_accuracy, future_revenue_accuracy)
            st.download_button(label="Download Excel Report",
                               data=output,
                               file_name=f"{filename}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # Complete the progress bar
            progress_bar.progress(100)

if __name__ == "__main__":
    main()
