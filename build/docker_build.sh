#!/bin/sh

time_now=$(date "+%Y%m%d")
git_commit_id=$(git rev-parse --short HEAD)
DOCKER_TAG_SUFFIX="release_"$time_now"_"$git_commit_id

IMAGE_NAME=word_parse:$DOCKER_TAG_SUFFIX

cd ..
docker build -t ${IMAGE_NAME} -f ./build/Dockerfile .
# docker push ${IMAGE_NAME}
# echo "push dev: ${IMAGE_NAME}"
