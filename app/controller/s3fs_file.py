#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os.path
import traceback
import urllib
from pathlib import Path

from loguru import logger
from minio import Minio
from minio.error import S3Error

from config.conf import S3_ENDPOINT, S3_KEY_ID, S3_KEY_SECRET, TEMP_PATH


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
        try:
            # Make the bucket if it doesn't exist.
            found = self.client.bucket_exists(bucket_name)
            if not found:
                # self.client.make_bucket(bucket_name)
                # logger.info(f"Created bucket {bucket_name}")
                logger.error(f"Bucket: {bucket_name} is not exists.")
                raise
            self.client.fput_object(bucket_name, obj_path, local_file)
            logger.info(f"file: {upload_path} upload success")
        except S3Error as exc:
            logger.error(f"error occurred.{exc}")

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
            logger.warning(f"Cannot delete {file_path}: file does not exist")

    def file_exists(self, file_path: str) -> bool:
        """
        Checks if a file exists in the MinIO bucket.
        
        :param file_path: The path of the file to check in the MinIO bucket.
        :return: True if the file exists, False otherwise.
        """
        bucket_name, obj_path = self.parse_bucket_name(file_path)
        try:
            self.client.stat_object(bucket_name, obj_path)
            logger.info(f"File {file_path} exists")
            return True
        except S3Error as exc:
            if exc.code == 'NoSuchKey':
                logger.info(f"File {file_path} does not exist")
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
        try:
            self.client.fget_object(bucket_name, object_name, download_path)
            logger.info(f"Downloaded {object_name} to {download_path} successfully")
        except S3Error as exc:
            logger.error(f"Error occurred while downloading: {exc}")

    def read_file(self, file_path):
        bucket_name, object_name = self.parse_bucket_name(file_path)
        try:
            data = self.client.get_object(bucket_name, object_name)
            return data.read().decode('utf-8')
        except S3Error as exc:
            print(traceback.print_exc())
            logger.error(f"Error occurred while read file, file_path: {file_path}, error: {exc}")
            return None

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
            logger.error(f"Error occurred while writing file to MinIO, file_path: {upload_path}, error: {exc}")
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)

    @staticmethod
    def parse_bucket_name(s3_path: str):
        path_obj = Path(s3_path)
        bucket_name = path_obj.parts[0]
        remaining_path = path_obj.relative_to(bucket_name)
        return bucket_name, remaining_path
    

s3_controller = MinioTool()
