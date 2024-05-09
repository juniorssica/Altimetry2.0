import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import os
import base64
import xlsxwriter


def plot_altimetry(data):
    # Convert distance from meters to kilometers and round to 2 decimal places
    data['Distance_km'] = (data['Distance'] / 1000).round(2)

    # Create a new column for interval in kilometers
    data['Interval'] = (data['Distance_km'] // 1).astype(int) * 1

    # Group data by kilometer intervals and calculate average altitudes
    data_grouped = data.groupby('Interval').mean()
    data_grouped = data_grouped[['Altitude', 'Distance_km']]
    data_grouped.rename(columns={'Altitude': 'Altitude_m'}, inplace=True)
    data_grouped = data_grouped[['Distance_km', 'Altitude_m']].round(2)

    return data_grouped


def get_excel_download_link(data, filename):
    """
    Function to get the download link for the Excel file.
    """
    excel_output_file = f'static/{filename}.xlsx'
    data.to_excel(excel_output_file, index=False)
    with pd.ExcelWriter(excel_output_file, engine='xlsxwriter') as writer:
        # Write the converted data to the first sheet
        data.to_excel(writer, sheet_name='profil_altimetry', index=False)

        # Create a workbook and add a worksheet for the Air3D chart
        workbook = writer.book
        worksheet = workbook.add_worksheet('Air3D_chart')

        # Create a chart object
        chart = workbook.add_chart({'type': 'scatter3D'})

        # Configure the chart
        chart.add_series({
            'name': ['profil_altimetry', 0, 1],
            'categories': ['profil_altimetry', 1, 0, len(data), 0],
            'values': ['profil_altimetry', 1, 1, len(data), 1],
            'marker': {'type': 'circle', 'size': 5},
        })

        # Add the chart to the worksheet
        worksheet.insert_chart('A1', chart)

        # Close the Pandas Excel writer and save the file
        writer.save()

    with open(excel_output_file, 'rb') as f:
        file_content = f.read()
    base64_encoded = base64.b64encode(file_content).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{base64_encoded}" download="{filename}.xlsx">Télécharger le fichier Excel</a>'
    return href

