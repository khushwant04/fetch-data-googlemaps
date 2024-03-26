import googlemaps
import requests
import pandas as pd
import openpyxl

# initialize Client 
api_key = 'API_KEY'


# initialize file
filename = 'data.xlsx'


# taking inputs from user and processing
address = input('Enter address:')
radius = int(input('Enter the radius:'))


# getting coordinates
def get_coordinates(address, api_key):
    url = f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key={api_key}"
    response = requests.get(url)
    data = response.json()
    if data['status'] == 'OK':
        location = data['results'][0]['geometry']['location']
        latitude = location['lat']
        longitude = location['lng']
        return latitude, longitude
    else:
        return None, None
      
# getting details of nearby locations    
def get_nearby_places(latitude,longitude,radius,api_key):
    url = f"https://maps.googleapis.com/maps/api/place/nearbysearch/json?location={latitude},{longitude}&radius={radius}&type={type}&key={api_key}"
    response = requests.get(url)
    data = response.json()
    return data    


# latitude and longitude output
latitude,longitude = get_coordinates(address,api_key)

# calling function for getting data
data = get_nearby_places(latitude,longitude,radius,api_key)

def save_to_excel():
    try:
        # Load existing data if file exists
        try:
            df_existing = pd.read_excel(filename)
        except FileNotFoundError:
            df_existing = pd.DataFrame()

        # Process new data
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
