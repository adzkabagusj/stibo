import os
import boto3

s3_client = boto3.client('s3')

# Nama Bucket (Tanpa path folder)
DESTINATION_BUCKET = 'lh-sadp-em-local-data-eu-west-1-wbeo'

def lambda_handler(event, context):
    try:
        # Menentukan path folder dan nama file di dalam bucket tujuan
        file_path = "bronze/inbound/map_frs/test-dummy-file.txt"
        dummy_content = "Halo! Testing koneksi lintas akun AWS."
        
        print(f"Mengirim ke: {DESTINATION_BUCKET}/{file_path}")
        
        s3_client.put_object(
            Bucket=DESTINATION_BUCKET,
            Key=file_path,
            Body=dummy_content.encode('utf-8')
        )
        
        return {
            'statusCode': 200,
            'body': f"Sukses! File terkirim ke {DESTINATION_BUCKET}/{file_path}"
        }
        
    except Exception as e:
        print(f"Error: {str(e)}")
        raise e