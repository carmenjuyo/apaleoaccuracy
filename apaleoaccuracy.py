import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.graph_objs as go
from plotly.subplots import make_subplots
from io import BytesIO

# Set page layout to wide
st.set_page_config(layout="wide", page_title="Guestline Daily Variance and Accuracy Calculator")

# Define the function to read from CSV
# Updated function to read CSV and support variance checks
def read_csv_data(file):
    df = pd.read_csv(file, delimiter=';', quotechar='"')

    # Clean up columns
    df.columns = [col.strip().replace('"', '') for col in df.columns]

    # Validate required columns
    required_columns = ['date', 'AF RNs', 'AF Rev']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns in the CSV: {', '.join(missing_columns)}")

    # Convert and clean up data
    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    if df['date'].isnull().any():
        raise ValueError("Some 'date' values could not be parsed. Ensure the column contains valid dates.")

    df['AF RNs'] = pd.to_numeric(df['AF RNs'], errors='coerce')
    df['AF Rev'] = pd.to_numeric(df['AF Rev'], errors='coerce')

    return df

# Function to read actuals from CSV
def read_actuals_csv(file):
    df = pd.read_csv(file, delimiter=';', quotechar='"')

    # Clean up columns
    df.columns = [col.strip().replace('"', '') for col in df.columns]

    # Validate required columns
    required_columns = ['arrivalDate', 'rn', 'revNet']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns in the CSV: {', '.join(missing_columns)}")

    # Convert and clean up data
    df['arrivalDate'] = pd.to_datetime(df['arrivalDate'], errors='coerce')
    if df['arrivalDate'].isnull().any():
        raise ValueError("Some 'arrivalDate' values could not be parsed. Ensure the column contains valid dates.")

    df['rn'] = pd.to_numeric(df['rn'], errors='coerce')
    df['revNet'] = pd.to_numeric(df['revNet'], errors='coerce')

    df = df.rename(columns={'arrivalDate': 'date', 'rn': 'soldCount', 'revNet': 'netAccommodationRevenue'})
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
            combined_df['Abs RN Accuracy'] = combined_df['Abs RN Accuracy'].str.rstrip('%').astype('float') / 100
            combined_df['Abs Rev Accuracy'] = combined_df['Abs Rev Accuracy'].str.rstrip('%').astype('float') / 100

            combined_df.to_excel(writer, sheet_name='Daily Variance Detail', index=False)
            worksheet_combined = writer.sheets['Daily Variance Detail']

            # Format columns
            format_number = workbook.add_format({'num_format': '#,##0.00'})  # Floats
            format_whole = workbook.add_format({'num_format': '0'})  # Whole numbers

            worksheet_combined.set_column('A:A', None, format_whole)  # Date
            worksheet_combined.set_column('B:B', None, format_whole)  # AF RNs
            worksheet_combined.set_column('C:C', None, format_number)  # AF Rev
            worksheet_combined.set_column('D:D', None, format_whole)  # SoldCount
            worksheet_combined.set_column('E:E', None, format_number)  # NetAccommodationRevenue
            worksheet_combined.set_column('F:F', None, format_whole)  # RN Diff
            worksheet_combined.set_column('G:G', None, format_number)  # Rev Diff
            worksheet_combined.set_column('H:H', None, format_percent)  # Abs RN Accuracy
            worksheet_combined.set_column('I:I', None, format_percent)  # Abs Rev Accuracy

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

    # File uploaders
    forecast_file = st.file_uploader("Upload Forecast File (.csv)", type=['csv'])
    actuals_file = st.file_uploader("Upload Actuals File (.csv)", type=['csv'])

    if forecast_file and actuals_file:
        try:
            # Read files
            forecast_data = read_csv_data(forecast_file)
            actuals_data = read_actuals_csv(actuals_file)

            # Perform accuracy check
            results_df = perform_accuracy_check(forecast_data, actuals_data)

            # Display results
            st.markdown("### Accuracy Results")
            styled_df = results_df.style.applymap(color_scale, subset=['RN Accuracy', 'Rev Accuracy'])
            st.dataframe(styled_df)

            # Calculate overall accuracies
            current_date = pd.to_datetime('today').normalize()
            past_mask = results_df['date'] < current_date
            future_mask = results_df['date'] >= current_date
            past_rooms_accuracy = (1 - (abs(results_df.loc[past_mask, 'RN Diff']).sum() / results_df.loc[past_mask, 'AF RNs'].sum())) * 100
            past_revenue_accuracy = (1 - (abs(results_df.loc[past_mask, 'Rev Diff']).sum() / results_df.loc[past_mask, 'AF Rev'].sum())) * 100
            future_rooms_accuracy = (1 - (abs(results_df.loc[future_mask, 'RN Diff']).sum() / results_df.loc[future_mask, 'AF RNs'].sum())) * 100
            future_revenue_accuracy = (1 - (abs(results_df.loc[future_mask, 'Rev Diff']).sum() / results_df.loc[future_mask, 'AF Rev'].sum())) * 100

            # Add Excel export functionality
            base_filename = "Accuracy_Report"
            excel_data, filename = create_excel_download(results_df, base_filename, past_rooms_accuracy, past_revenue_accuracy, future_rooms_accuracy, future_revenue_accuracy)
            st.download_button(
                label="Download Results as Excel",
                data=excel_data,
                file_name=f"{filename}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing files: {e}")

if __name__ == "__main__":
    main()
