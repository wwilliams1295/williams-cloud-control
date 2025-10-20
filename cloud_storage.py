#!/usr/bin/env python3
"""
Cloud Storage Adapter for Render Deployment
Handles file storage in cloud services for production deployment.
"""

import os
import logging
from typing import Dict, Any, Optional, BinaryIO
from pathlib import Path
import tempfile
import asyncio

logger = logging.getLogger(__name__)

class CloudStorageAdapter:
    """Handles file storage for cloud deployment."""
    
    def __init__(self):
        self.storage_type = os.getenv("STORAGE_TYPE", "local")  # local, s3, gcs, azure
        self.bucket_name = os.getenv("STORAGE_BUCKET", "")
        self.region = os.getenv("STORAGE_REGION", "us-east-1")
        
        # Initialize storage client based on type
        self.client = self._init_storage_client()
        
        # Base URL for public access
        self.public_base_url = os.getenv("PUBLIC_BASE_URL", "http://127.0.0.1:8000")
        self.storage_base_url = os.getenv("STORAGE_BASE_URL", "")
    
    def _init_storage_client(self):
        """Initialize the appropriate storage client."""
        if self.storage_type == "s3":
            return self._init_s3_client()
        elif self.storage_type == "gcs":
            return self._init_gcs_client()
        elif self.storage_type == "azure":
            return self._init_azure_client()
        else:
            return self._init_local_client()
    
    def _init_s3_client(self):
        """Initialize AWS S3 client."""
        try:
            import boto3
            return boto3.client(
                's3',
                aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
                aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
                region_name=self.region
            )
        except ImportError:
            logger.error("boto3 not installed. Install with: pip install boto3")
            return None
        except Exception as e:
            logger.error(f"Error initializing S3 client: {e}")
            return None
    
    def _init_gcs_client(self):
        """Initialize Google Cloud Storage client."""
        try:
            from google.cloud import storage
            return storage.Client()
        except ImportError:
            logger.error("google-cloud-storage not installed. Install with: pip install google-cloud-storage")
            return None
        except Exception as e:
            logger.error(f"Error initializing GCS client: {e}")
            return None
    
    def _init_azure_client(self):
        """Initialize Azure Blob Storage client."""
        try:
            from azure.storage.blob import BlobServiceClient
            connection_string = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
            return BlobServiceClient.from_connection_string(connection_string)
        except ImportError:
            logger.error("azure-storage-blob not installed. Install with: pip install azure-storage-blob")
            return None
        except Exception as e:
            logger.error(f"Error initializing Azure client: {e}")
            return None
    
    def _init_local_client(self):
        """Initialize local file system client."""
        # Create local storage directory
        local_dir = Path("/tmp/jarvis_files")
        local_dir.mkdir(parents=True, exist_ok=True)
        return {"type": "local", "path": local_dir}
    
    async def save_file(self, file_path: str, content: bytes, content_type: str = None) -> Dict[str, Any]:
        """Save a file to cloud storage."""
        try:
            filename = Path(file_path).name
            
            if self.storage_type == "s3":
                return await self._save_to_s3(filename, content, content_type)
            elif self.storage_type == "gcs":
                return await self._save_to_gcs(filename, content, content_type)
            elif self.storage_type == "azure":
                return await self._save_to_azure(filename, content, content_type)
            else:
                return await self._save_to_local(filename, content)
                
        except Exception as e:
            logger.error(f"Error saving file {file_path}: {e}")
            return {"success": False, "error": str(e)}
    
    async def _save_to_s3(self, filename: str, content: bytes, content_type: str = None) -> Dict[str, Any]:
        """Save file to AWS S3."""
        if not self.client:
            return {"success": False, "error": "S3 client not initialized"}
        
        try:
            key = f"jarvis-files/{filename}"
            self.client.put_object(
                Bucket=self.bucket_name,
                Key=key,
                Body=content,
                ContentType=content_type or "application/octet-stream"
            )
            
            public_url = f"https://{self.bucket_name}.s3.{self.region}.amazonaws.com/{key}"
            
            return {
                "success": True,
                "filename": filename,
                "public_url": public_url,
                "storage_path": key
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _save_to_gcs(self, filename: str, content: bytes, content_type: str = None) -> Dict[str, Any]:
        """Save file to Google Cloud Storage."""
        if not self.client:
            return {"success": False, "error": "GCS client not initialized"}
        
        try:
            bucket = self.client.bucket(self.bucket_name)
            blob = bucket.blob(f"jarvis-files/{filename}")
            
            blob.upload_from_string(content, content_type=content_type or "application/octet-stream")
            blob.make_public()
            
            public_url = blob.public_url
            
            return {
                "success": True,
                "filename": filename,
                "public_url": public_url,
                "storage_path": f"jarvis-files/{filename}"
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _save_to_azure(self, filename: str, content: bytes, content_type: str = None) -> Dict[str, Any]:
        """Save file to Azure Blob Storage."""
        if not self.client:
            return {"success": False, "error": "Azure client not initialized"}
        
        try:
            container_name = "jarvis-files"
            blob_name = filename
            
            blob_client = self.client.get_blob_client(container=container_name, blob=blob_name)
            blob_client.upload_blob(content, content_type=content_type or "application/octet-stream")
            
            public_url = blob_client.url
            
            return {
                "success": True,
                "filename": filename,
                "public_url": public_url,
                "storage_path": f"{container_name}/{blob_name}"
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _save_to_local(self, filename: str, content: bytes) -> Dict[str, Any]:
        """Save file to local storage (for development)."""
        try:
            local_path = self.client["path"] / filename
            local_path.write_bytes(content)
            
            public_url = f"{self.public_base_url}/files/{filename}"
            
            return {
                "success": True,
                "filename": filename,
                "public_url": public_url,
                "storage_path": str(local_path)
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def get_file(self, filename: str) -> Dict[str, Any]:
        """Retrieve a file from cloud storage."""
        try:
            if self.storage_type == "s3":
                return await self._get_from_s3(filename)
            elif self.storage_type == "gcs":
                return await self._get_from_gcs(filename)
            elif self.storage_type == "azure":
                return await self._get_from_azure(filename)
            else:
                return await self._get_from_local(filename)
        except Exception as e:
            logger.error(f"Error retrieving file {filename}: {e}")
            return {"success": False, "error": str(e)}
    
    async def _get_from_s3(self, filename: str) -> Dict[str, Any]:
        """Get file from S3."""
        if not self.client:
            return {"success": False, "error": "S3 client not initialized"}
        
        try:
            key = f"jarvis-files/{filename}"
            response = self.client.get_object(Bucket=self.bucket_name, Key=key)
            content = response['Body'].read()
            
            return {
                "success": True,
                "content": content,
                "content_type": response.get('ContentType', 'application/octet-stream')
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _get_from_gcs(self, filename: str) -> Dict[str, Any]:
        """Get file from GCS."""
        if not self.client:
            return {"success": False, "error": "GCS client not initialized"}
        
        try:
            bucket = self.client.bucket(self.bucket_name)
            blob = bucket.blob(f"jarvis-files/{filename}")
            content = blob.download_as_bytes()
            
            return {
                "success": True,
                "content": content,
                "content_type": blob.content_type or "application/octet-stream"
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _get_from_azure(self, filename: str) -> Dict[str, Any]:
        """Get file from Azure."""
        if not self.client:
            return {"success": False, "error": "Azure client not initialized"}
        
        try:
            container_name = "jarvis-files"
            blob_client = self.client.get_blob_client(container=container_name, blob=filename)
            content = blob_client.download_blob().readall()
            
            return {
                "success": True,
                "content": content,
                "content_type": blob_client.get_blob_properties().content_settings.content_type or "application/octet-stream"
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _get_from_local(self, filename: str) -> Dict[str, Any]:
        """Get file from local storage."""
        try:
            local_path = self.client["path"] / filename
            if not local_path.exists():
                return {"success": False, "error": "File not found"}
            
            content = local_path.read_bytes()
            
            return {
                "success": True,
                "content": content,
                "content_type": "application/octet-stream"
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def get_public_url(self, filename: str) -> str:
        """Get public URL for a file."""
        if self.storage_type == "s3":
            return f"https://{self.bucket_name}.s3.{self.region}.amazonaws.com/jarvis-files/{filename}"
        elif self.storage_type == "gcs":
            return f"https://storage.googleapis.com/{self.bucket_name}/jarvis-files/{filename}"
        elif self.storage_type == "azure":
            return f"https://{self.bucket_name}.blob.core.windows.net/jarvis-files/{filename}"
        else:
            return f"{self.public_base_url}/files/{filename}"

# Global storage adapter
storage_adapter = CloudStorageAdapter()

async def demo_cloud_storage():
    """Demonstrate cloud storage functionality."""
    print("☁️ Cloud Storage Demo")
    print("=" * 50)
    
    # Test file content
    test_content = b"Hello, this is a test file for cloud storage!"
    
    # Save file
    print("\n1. Saving file to cloud storage...")
    result = await storage_adapter.save_file("test.txt", test_content, "text/plain")
    print(f"Result: {result}")
    
    if result.get("success"):
        print(f"Public URL: {result['public_url']}")
        
        # Retrieve file
        print("\n2. Retrieving file from cloud storage...")
        get_result = await storage_adapter.get_file("test.txt")
        print(f"Retrieved: {get_result.get('success', False)}")
        if get_result.get("success"):
            print(f"Content: {get_result['content'].decode()[:50]}...")

if __name__ == "__main__":
    asyncio.run(demo_cloud_storage())
