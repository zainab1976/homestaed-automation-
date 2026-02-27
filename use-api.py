# Example: How to use your Medical Assessment API with Python

import requests
import json

API_BASE = 'http://localhost:3000'

def use_api():
    print('🚀 Using Medical Assessment API...\n')

    try:
        # 1. Check API health
        print('1. Checking API health...')
        health_response = requests.get(f'{API_BASE}/health')
        health_data = health_response.json()
        print(f'✅ API Status: {health_data["status"]}')
        print(f'📅 Timestamp: {health_data["timestamp"]}')

        # 2. List available files
        print('\n2. Listing available files...')
        files_response = requests.get(f'{API_BASE}/api/files')
        files_data = files_response.json()
        print('📁 Available files:')
        for file in files_data['files']:
            print(f'   - {file["name"]} ({file["type"]})')

        # 3. Read Excel file data
        print('\n3. Reading Excel file data...')
        excel_response = requests.get(f'{API_BASE}/api/excel/FMG%2009.22.2025.xlsx')
        excel_data = excel_response.json()
        
        if excel_data['success']:
            print('📊 Excel file sheets:')
            for sheet_name, data in excel_data['data'].items():
                print(f'   - {sheet_name}: {len(data)} rows')

        # 4. Process assessments (uncomment to run automation)
        print('\n4. Ready to process assessments...')
        print('⚠️  To start processing, uncomment the code below:')
        print('''
        process_response = requests.post(
            f'{API_BASE}/api/process/FMG%2009.22.2025.xlsx',
            headers={'Content-Type': 'application/json'},
            json={'headless': True, 'slowMo': 2000}
        )
        process_data = process_response.json()
        print('🔄 Processing started:', process_data)
        ''')

    except Exception as error:
        print(f'❌ Error: {error}')

if __name__ == '__main__':
    use_api()




