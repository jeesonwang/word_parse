#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os.path
import urllib
from pathlib import Path

from loguru import logger
from minio import Minio
from minio.error import S3Error, MinioException

from config.conf import S3_ENDPOINT, S3_KEY_ID, S3_KEY_SECRET, TEMP_PATH, FS_TYPE


class MinioTool():
    def __init__(self):
        self.client = Minio(endpoint=urllib.parse.urlparse(S3_ENDPOINT).netloc,
                     access_key=S3_KEY_ID,
                     secret_key=S3_KEY_SECRET,
                     secure=False)

    def upload(self, local_file, upload_path):
        """
        :param local_file: 本地文件
        :param upload_path: s3文件路径
        :return:
        """
        bucket_name, obj_path = self.parse_bucket_name(upload_path)
        if self.file_exists(upload_path):
            raise MinioException(f"{upload_path} file is exist in S3. Please do not upload repeatedly.")
        self.client.fput_object(bucket_name, obj_path, local_file)
        logger.info(f"file: {upload_path} upload success")

    def delete_file(self, file_path):
        """
        Deletes a file from the MinIO bucket.
        
        :param file_path: The path of the file to delete in the MinIO bucket.
        :return: None
        """
        bucket_name, object_name = self.parse_bucket_name(file_path)
        if self.file_exists(file_path):
            self.client.remove_object(bucket_name, object_name)
            logger.info(f"Deleted {file_path} successfully")
        else:
            logger.warning(f"Cannot delete {file_path}")

    def file_exists(self, file_path: str) -> bool:
        """
        Checks if a file exists in the MinIO bucket.
        
        :param file_path: The path of the file to check in the MinIO bucket.
        :return: True if the file exists, False otherwise.
        """
        bucket_name, obj_path = self.parse_bucket_name(file_path)
        try:
            self.client.stat_object(bucket_name, obj_path)
            return True
        except S3Error as exc:
            if exc.code == 'NoSuchKey':
                return False
            logger.error(f"Error occurred while checking if file exists: {exc}")
            return False

    def download_file(self, file_path, download_path):
        """
        Downloads a file from the MinIO bucket to a local path.
        
        :param file_path: The path of the file in the MinIO.
        :param download_path: The local path where the file will be downloaded.
        :return: None
        """
        bucket_name, object_name = self.parse_bucket_name(file_path)
        self.client.fget_object(bucket_name, object_name, download_path)
        
    def read_file(self, file_path):
        bucket_name, object_name = self.parse_bucket_name(file_path)
        data = self.client.get_object(bucket_name, object_name)
        return data.read().decode('utf-8')

    def write_file(self, upload_path: str, content: str):
        """
        将字符串内容写入 MinIO 中指定路径的文件。
        
        :param upload_path: MinIO 中的目标文件路径。
        :param content: 要写入的字符串内容。
        :return: None
        """
        # 创建临时文件以便上传到 MinIO
        temp_file = os.path.join(TEMP_PATH, 'temp_write_file.txt')
        try:
            with open(temp_file, 'w', encoding='utf-8') as f:
                f.write(content)

            self.upload(temp_file, upload_path)
            logger.info(f"File {upload_path} successfully written to MinIO at path {upload_path}")

        except S3Error as exc:
            raise MinioException(f"Error occurred while writing file to MinIO, file_path: {upload_path}, error: {exc}")
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)

    @staticmethod
    def parse_bucket_name(s3_path: str):
        path_obj = Path(s3_path)
        bucket_name = path_obj.parts[0]
        remaining_path = path_obj.relative_to(bucket_name)
        return str(bucket_name), str(remaining_path)


S3_controllers = {
    "minio": MinioTool,
    }

s3_controller = S3_controllers[FS_TYPE]()
