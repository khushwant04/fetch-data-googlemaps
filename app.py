import openpyxl
from flask import Flask, request, jsonify
import json
import requests
import pandas as pd

api_key = 'api'
filename = 'data.xlsx'

app = Flask(__name__)


@app.route('/',methods=['POST'])
def save_to_excel():
    try:
        # Load existing data if file exists
        try:
            df_existing = pd.read_excel(filename)
        except FileNotFoundError:
            df_existing = pd.DataFrame()

        # Process new data
        data = request.json
        locations = []
        for place in data['places']:
            types = place['types']
            formatted_address = place['formattedAddress']
            website_uri = place['websiteUri']
            display_name = place['displayName']['text']
            locations.append((display_name, types, formatted_address, website_uri))

        # Append new data to existing data
        df_new = pd.DataFrame(locations, columns=['Display Name', 'Types', 'Formatted Address', 'Website'])
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)

        # Save the combined data to Excel
        df_combined.to_excel(filename, index=False)
        
        return 'Data appended and saved successfully'
    
    except Exception as e:
        return f'Error: {str(e)}'

if __name__ == '__main__':
    app.run(debug=True)