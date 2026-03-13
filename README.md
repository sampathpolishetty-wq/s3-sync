import json
import boto3
import requests
import os
import io
from msal import ConfidentialClientApplication
from datetime import datetime, timezone
import pandas as pd
import openpyxl

s3_client = boto3.client('s3')

def lambda_handler(event, context):
    """
    Complete SharePoint to S3 sync with Excel processing
    Trigger with event: {"folder": "/Shared Documents/Reports"}
    """
    # Environment variables (set in Lambda console)
    sharepoint_site = os.environ['SHAREPOINT_SITE']  # 'https://tenant.sharepoint.com/sites/yoursite'
    client_id = os.environ['CLIENT_ID']
    client_secret = os.environ['CLIENT_SECRET']
    tenant_id = os.environ['TENANT_ID']
    bucket_name = os.environ['S3_BUCKET']
    s3_prefix = os.environ.get('S3_PREFIX', 'sharepoint-files/')
    
    # Event parameters (optional)
    source_folder = event.get('folder', '/Shared Documents')
    file_filter = event.get('file_filter', ['.xlsx', '.xls', '.csv', '.pdf'])  # Extensions to process
    process_excel = event.get('process_excel', True)
    
    print(f"Scanning {sharepoint_site}{source_folder}")
    
    # Get Microsoft Graph access token
    access_token = get_access_token(client_id, client_secret, tenant_id)
    headers = {'Authorization': f'Bearer {access_token}'}
    
    # Scan and process files
    uploaded_files = []
    files = scan_sharepoint_folder(sharepoint_site, source_folder, headers, file_filter)
    
    for file_info in files:
        try:
            success = process_and_upload_file(
                sharepoint_site, file_info, headers, bucket_name, s3_prefix, process_excel
            )
            if success:
                uploaded_files.append(file_info['relative_path'])
        except Exception as e:
            print(f"Failed to process {file_info['name']}: {str(e)}")
    
    return {
        'statusCode': 200,
        'body': json.dumps({
            'message': f'Processed {len(uploaded_files)} files',
            'uploaded': uploaded_files,
            'total_scanned': len(files)
        })
    }

def get_access_token(client_id, client_secret, tenant_id):
    """Acquire Microsoft Graph API access token"""
    app = ConfidentialClientApplication(
        client_id, 
        authority=f'https://login.microsoftonline.com/{tenant_id}',
        client_credential=client_secret,
        scopes=['https://graph.microsoft.com/.default']
    )
    token_result = app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])
    if 'access_token' not in token_result:
        raise Exception(f"Token acquisition failed: {token_result.get('error_description')}")
    return token_result['access_token']

def scan_sharepoint_folder(site_url, folder_path, headers, file_filter, relative_path='', all_files=None):
    """Recursively scan SharePoint folder for files matching filter"""
    if all_files is None:
        all_files = []
    
    # List children with pagination support
    children_url = f"{site_url}{folder_path}:/children?$select=id,name,size,lastModifiedDateTime,file,folder,parentReference&$top=1000"
    while children_url:
        resp = requests.get(children_url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        children = data.get('value', [])
        next_link = data.get('@odata.nextLink')
        children_url = next_link if next_link else None
        
        for item in children:
            full_rel_path = f"{relative_path}/{item['name']}" if relative_path else item['name']
            item_path = item['parentReference']['path'] + '/' + item['name']
            
            # Check file extension filter
            name_lower = item['name'].lower()
            if any(name_lower.endswith(ext) for ext in file_filter) and 'file' in item:
                all_files.append({
                    'name': item['name'],
                    'path': item_path,
                    'relative_path': full_rel_path,
                    'size': item['size'],
                    'lastModifiedDateTime': item['lastModifiedDateTime']
                })
            elif 'folder' in item:
                # Recurse into subfolders
                scan_sharepoint_folder(site_url, item_path, headers, file_filter, full_rel_path, all_files)
    
    return all_files

def process_excel_file(content, file_info):
    """Convert Excel to cleaned CSV"""
    try:
        df = pd.read_excel(io.BytesIO(content))
        
        # Data cleaning
        df = df.dropna(how='all').dropna(axis=1, how='all')
        df.columns = [str(col).strip().lower().replace(' ', '_') for col in df.columns]
        
        # Convert to CSV
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        processed_content = csv_buffer.getvalue().encode('utf-8')
        
        metadata = {
            'original_rows': len(df),
            'processed_rows': len(df.dropna()),
            'columns': list(df.columns),
            'source_last_modified': file_info['lastModifiedDateTime']
        }
        
        return processed_content, metadata, 'text/csv'
    
    except Exception as e:
        print(f"Excel processing failed: {e}")
        return content, {}, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

def process_and_upload_file(site_url, file_info, headers, bucket_name, s3_prefix, process_excel):
    """Download file from SharePoint, process if Excel, upload to S3"""
    content_url = f"{site_url}{file_info['path']}:content"
    
    # Download file
    resp = requests.get(content_url, headers=headers, stream=True)
    resp.raise_for_status()
    raw_content = resp.content
    
    # Process Excel files
    content_type = resp.headers.get('content-type', 'application/octet-stream')
    metadata = {'source_path': file_info['path']}
    
    file_lower = file_info['name'].lower()
    if process_excel and any(file_lower.endswith(ext) for ext in ['.xlsx', '.xls', '.xlsm']):
        processed_content, excel_metadata, csv_content_type = process_excel_file(raw_content, file_info)
        s3_key = f"{s3_prefix}{file_info['relative_path']}.csv"
        content_type = csv_content_type
        metadata.update(excel_metadata)
    else:
        processed_content = raw_content
        s3_key = f"{s3_prefix}{file_info['relative_path']}"
    
    # Upload to S3
    s3_client.put_object(
        Bucket=bucket_name,
        Key=s3_key,
        Body=processed_content,
        ContentType=content_type,
        Metadata=metadata
    )
    
    print(f"Uploaded {file_info['name']} -> {s3_key} ({len(processed_content)} bytes)")
    return True
