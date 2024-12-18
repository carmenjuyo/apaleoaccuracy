import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.graph_objs as go
from plotly.subplots import make_subplots
from io import BytesIO

# Set page layout to wide
st.set_page_config(layout="wide", page_title="Apaleo Daily Variance and Accuracy Calculator")

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
        expected_columns = ['arrivalDate', 'rn', 'revNet', 'revTotal', 'revFb', 'revResTotal']
        df = pd.read_csv(file, delimiter=';', quotechar='"')
        for col in expected_columns:
            if col not in df.columns:
                raise ValueError(f"Expected column '{col}' not found in the uploaded file.")
        df = df[['arrivalDate', 'rn', 'revNet']]
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
    # Center the title
    st.markdown("<h1 style='text-align: center;'> Guestline Daily Variance and Accuracy Calculator</h1>", unsafe_allow_html=True)

    # File uploaders
    col1, col2 = st.columns(2)
    with col1:
        history_forecast_file = st.file_uploader("Upload History and Forecast (.csv)", type=['csv'])
    with col2:
        daily_totals_file = st.file_uploader("Upload Daily Totals (.csv)", type=['csv'])

    if history_forecast_file and daily_totals_file:
        if st.button("Process Data"):
            try:
                progress_bar = st.progress(0)

                # Read the files
                hf_df = read_data(history_forecast_file, is_history_and_forecast_file=True)
                progress_bar.progress(25)

                dt_df = read_data(daily_totals_file, is_history_and_forecast_file=False)
                progress_bar.progress(50)

                # Merge data
                merged_df = pd.merge(hf_df, dt_df, on='date', how='inner')
                progress_bar.progress(60)

                # Calculate discrepancies
                merged_df['RN Diff'] = merged_df['Juyo RN'] - merged_df['AF RNs']
                merged_df['Rev Diff'] = merged_df['Juyo Rev'] - merged_df['AF Rev']

                # Calculate absolute accuracy
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

                # Format accuracy percentages
                merged_df['Abs RN Accuracy'] = merged_df['Abs RN Accuracy'].map(lambda x: f"{x:.2f}%")
                merged_df['Abs Rev Accuracy'] = merged_df['Abs Rev Accuracy'].map(lambda x: f"{x:.2f}%")

                progress_bar.progress(75)

                # Calculate overall accuracies
                current_date = pd.to_datetime('today').date()
                past_mask = merged_df['date'] < current_date
                future_mask = merged_df['date'] >= current_date
                past_rooms_accuracy = (1 - abs(merged_df.loc[past_mask, 'RN Diff']).sum() / merged_df.loc[past_mask, 'AF RNs'].sum()) * 100
                past_revenue_accuracy = (1 - abs(merged_df.loc[past_mask, 'Rev Diff']).sum() / merged_df.loc[past_mask, 'AF Rev'].sum()) * 100
                future_rooms_accuracy = (1 - abs(merged_df.loc[future_mask, 'RN Diff']).sum() / merged_df.loc[future_mask, 'AF RNs'].sum()) * 100
                future_revenue_accuracy = (1 - abs(merged_df.loc[future_mask, 'Rev Diff']).sum() / merged_df.loc[future_mask, 'AF Rev'].sum()) * 100

                # Accuracy Matrix Table
                st.markdown("### Accuracy Matrix")
                accuracy_data = {
                    "Metric": ["RNs", "Revenue"],
                    "Past": [f"{past_rooms_accuracy:.2f}%", f"{past_revenue_accuracy:.2f}%"],
                    "Future": [f"{future_rooms_accuracy:.2f}%", f"{future_revenue_accuracy:.2f}%"]
                }
                accuracy_df = pd.DataFrame(accuracy_data)
                st.table(accuracy_df)

                # Day-by-Day Comparison
                st.markdown("### Day-by-Day Comparison")
                styled_df = merged_df.style.applymap(
                    lambda val: 'background-color: #469798; color: white' if isinstance(val, str) and val.endswith('%') and float(val.strip('%')) >= 98 else
                                'background-color: #F2A541; color: white' if isinstance(val, str) and val.endswith('%') and 95 <= float(val.strip('%')) < 98 else
                                'background-color: #BF3100; color: white',
                    subset=['Abs RN Accuracy', 'Abs Rev Accuracy']
                )
                st.dataframe(styled_df)

                # Visualization
                st.markdown("### Visualizations")
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

                # Update layout
                fig.update_layout(
                    height=600,
                    title='Discrepancies Over Time',
                    xaxis_title='Date',
                    yaxis_title='RN Discrepancy',
                    yaxis2_title='Revenue Discrepancy'
                )

                st.plotly_chart(fig, use_container_width=True)

                progress_bar.progress(90)

                # Define Excel export function
                def create_excel_download():
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        merged_df.to_excel(writer, index=False, sheet_name='Variance')
                    output.seek(0)
                    return output

                # Add download button
                excel_data = create_excel_download()
                st.download_button(label="Download Excel Report", data=excel_data, file_name="Variance_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                progress_bar.progress(100)

            except Exception as e:
                st.error(f"Error processing files: {e}")

if __name__ == "__main__":
    main()
