FROM swr.cn-north-4.myhuaweicloud.com/ddn-k8s/docker.io/openjdk:8-jdk

ENV TZ=Asia/Shanghai
ENV LANG=en_US.UTF-8

RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ >/etc/timezone \
    && sed -i 's#http://deb.debian.org#https://mirrors.163.com#g' /etc/apt/sources.list \
    && apt-get update \
    && apt-get -y install python3 python3-pip \
    && mkdir -p /project/word_parse

WORKDIR /project/word_parse

ADD requirements.txt /project/word_parse

RUN pip3 install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com

COPY . /project/word_parse/

EXPOSE 5000

CMD ["gunicorn", "-c", "gunicorn.conf.py", "server:create_app()", "–access-logfile -"]



