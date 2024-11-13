from config.conf import DEBUG

# 绑定端口 注意0.0.0.0的配置在docker网络中是关键配置
bind = "0.0.0.0:5000"
# 工作进程数 常规使用Docker横向扩容，故按项目实际情况配置
workers = 2
# 允许pending状态最大连接数 推荐64-2048
backlog = 2048
# 超时时间 根据实际情况配置 推荐30/60
# timeout = 60
# 是否使用debug模式
debug = DEBUG
# 是否重定向错误到日志文件
capture_output = True
# 工作单位 除特殊情况外,工程无脑使用gevent 算法可考虑使用gthread
# 详见文档: https://docs.gunicorn.org/en/stable/design.html#choosing-a-worker-type
# 不使用sync的最大问题在于sync虽然在短连接中网络模式最为简单干净,但不支持长连接,在高并发任务中的TCP效率会非常低
worker_class = "gevent"
# 每执行多少请求,即重启服务 该功能用于防止内存泄漏
# 对算法来说,该功能会导致重新初始化服务 可以设置0为关闭
max_requests = 50000
# 在max_requests的基础上,随机加减max_requests_jitter的参数值 该功能用于防止并发下所有容器同时重启
# 当max_requests=0时,该配置不生效
max_requests_jitter = 1000
# 客户端最大同时连接数 在极高并发根据服务器和实例数根据情况修改
# 仅适用于eventlet/gevent
worker_connections = 5000
# server端保持连接时间(秒) 根据情况设置2-5
keepalive = 60
