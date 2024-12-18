import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.graph_objs as go
import pytz
from plotly.subplots import make_subplots
from io import BytesIO

# Set page layout to wide
st.set_page_config(layout="wide", page_title="Apaleo Daily Variance and Accuracy Calculator")

# Define the function to read the two CSV files
def read_data(file, is_history_and_forecast_file, revenue_type=None):
    if not file.name.endswith('.csv'):
        raise ValueError("Unsupported file format. Please upload a .csv file.")

    if is_history_and_forecast_file:
        # History and Forecast file
        expected_columns = ['businessDay', 'soldCount', 'noShowsCount', 'netAccommodationRevenue', 'netRevenue', 'grossRevenue']
        df = pd.read_csv(file)
        for col in expected_columns:
            if col not in df.columns:
                raise ValueError(f"Expected column '{col}' not found in the uploaded file.")
        
        # Choose the revenue column dynamically
        revenue_column = revenue_type if revenue_type in df.columns else 'netAccommodationRevenue'
        df = df[['businessDay', 'soldCount', revenue_column]]
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
def create_excel_download(merged_df, accuracy_data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # --- Write the Accuracy Matrix ---
        # Ensure accuracy_data values are in decimals, not strings
        accuracy_df = pd.DataFrame({
            "Metric": accuracy_data["Metric"],
            "Past": [float(val.strip('%')) / 100 for val in accuracy_data["Past"]],
            "Future": [float(val.strip('%')) / 100 for val in accuracy_data["Future"]]
        })

        accuracy_df.to_excel(writer, sheet_name='Accuracy Matrix', index=False, startrow=1)
        worksheet_accuracy = writer.sheets['Accuracy Matrix']

        # Define formats for Accuracy Matrix
        format_green = workbook.add_format({'bg_color': '#469798', 'font_color': 'white'})
        format_yellow = workbook.add_format({'bg_color': '#F2A541', 'font_color': 'white'})
        format_red = workbook.add_format({'bg_color': '#BF3100', 'font_color': 'white'})
        format_percent = workbook.add_format({'num_format': '0.00%'})

        # Apply formatting to Accuracy Matrix columns
        worksheet_accuracy.set_column('B:C', None, format_percent)

        # Apply conditional formatting to the Accuracy Matrix
        worksheet_accuracy.conditional_format('B3:C4', {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})
        worksheet_accuracy.conditional_format('B3:C4', {'type': 'cell', 'criteria': 'between', 'minimum': 0.95, 'maximum': 0.98, 'format': format_yellow})
        worksheet_accuracy.conditional_format('B3:C4', {'type': 'cell', 'criteria': '<', 'value': 0.95, 'format': format_red})

        # --- Write the Variance Details ---
        # Convert accuracy percentages back to decimals for proper formatting
        temp_df = merged_df.copy()
        temp_df['Abs RN Accuracy'] = temp_df['Abs RN Accuracy'].str.rstrip('%').astype(float) / 100
        temp_df['Abs Rev Accuracy'] = temp_df['Abs Rev Accuracy'].str.rstrip('%').astype(float) / 100

        temp_df.to_excel(writer, sheet_name='Variance Detail', index=False)
        worksheet_variance = writer.sheets['Variance Detail']

        # Apply formats to Variance Detail
        worksheet_variance.set_column('H:I', None, format_percent)  # Accuracy columns
        worksheet_variance.conditional_format('H2:H{}'.format(len(temp_df) + 1),
                                              {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})
        worksheet_variance.conditional_format('H2:H{}'.format(len(temp_df) + 1),
                                              {'type': 'cell', 'criteria': 'between', 'minimum': 0.95, 'maximum': 0.98, 'format': format_yellow})
        worksheet_variance.conditional_format('H2:H{}'.format(len(temp_df) + 1),
                                              {'type': 'cell', 'criteria': '<', 'value': 0.95, 'format': format_red})
        worksheet_variance.conditional_format('I2:I{}'.format(len(temp_df) + 1),
                                              {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})
        worksheet_variance.conditional_format('I2:I{}'.format(len(temp_df) + 1),
                                              {'type': 'cell', 'criteria': 'between', 'minimum': 0.95, 'maximum': 0.98, 'format': format_yellow})
        worksheet_variance.conditional_format('I2:I{}'.format(len(temp_df) + 1),
                                              {'type': 'cell', 'criteria': '<', 'value': 0.95, 'format': format_red})

    output.seek(0)
    return output

def main():
    # Center the title
    st.markdown("<h1 style='text-align: center;'> Guestline Daily Variance and Accuracy Calculator</h1>", unsafe_allow_html=True)

    # File uploaders
    col1, col2 = st.columns(2)
    with col1:
        history_forecast_file = st.file_uploader("Upload History and Forecast (.csv)", type=['csv'])
    with col2:
        daily_totals_file = st.file_uploader("Upload Daily Totals (.csv)", type=['csv'])

    # Revenue Type Selection
    revenue_type = None
    if history_forecast_file:
        revenue_type = st.radio("Select Revenue Type for History and Forecast", ["netAccommodationRevenue", "grossRevenue"], index=0)

    if history_forecast_file and daily_totals_file:
        if st.button("Process Data"):
            try:
                progress_bar = st.progress(0)

                # Read files with dynamic revenue type
                hf_df = read_data(history_forecast_file, is_history_and_forecast_file=True, revenue_type=revenue_type)
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

                styled_accuracy_df = accuracy_df.style.applymap(
                    lambda val: 'background-color: #469798; color: white;' if float(val.strip('%')) >= 98 else
                                'background-color: #F2A541; color: white;' if 95 <= float(val.strip('%')) < 98 else
                                'background-color: #BF3100; color: white;',
                    subset=['Past', 'Future']
                )
                st.table(styled_accuracy_df)

                # Visualization and Day-by-Day Comparison
                st.markdown("### Day-by-Day Comparison")
                st.dataframe(merged_df)

                # Excel download
                excel_filename = f"{daily_totals_file.name.split('_')[0]}_ACCURACY_REPORT_{datetime.now().strftime('%Y%m%d_%H%M%S')}_CET.xlsx"
                excel_data = create_excel_download(merged_df, accuracy_data)
                st.download_button(label="Download Excel Report", data=excel_data, file_name=excel_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                progress_bar.progress(100)

            except Exception as e:
                st.error(f"Error processing files: {e}")

