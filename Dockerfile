FROM python:3.10-slim

# 设置国内源
RUN pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple

# 创建工作目录
WORKDIR /app

# 拷贝 requirements.txt 提前安装依赖
COPY requirements.txt ./

# 使用 requirements.txt 安装依赖
RUN pip install -r requirements.txt

# 拷贝其他项目文件
COPY . .

# 默认端口
EXPOSE 8001

# 启动 MCP 服务
CMD ["python", "main.py"]
