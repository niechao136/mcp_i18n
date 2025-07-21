import json
import os
import tempfile
from datetime import datetime, timezone

import pandas as pd
import requests
from fastapi.staticfiles import StaticFiles
from mcp.server.fastmcp import FastMCP

API_URL = "https://genieai.wise-apps.com:18081/v1/chat-messages"
API_KEY = "app-Sb0viPbp1QnIAdk3lKgtacEK"


def call_chatflow_with_markdown(markdown_table: str) -> str:
    """调用 Chatflow 接口处理 Markdown 表格"""
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "inputs": {
            "file_md": markdown_table
        },
        "query": "翻译i18文件",
        "response_mode": "blocking",
        "conversation_id": "",
        "user": "test",
        "files": [],
    }
    response = requests.post(API_URL, json=payload, headers=headers)
    response.raise_for_status()
    result = response.json()

    # 提取 answer 字符串并解析成 JSON
    answer_str = result.get("answer", "")

    try:
        cleaned = answer_str.encode('unicode_escape').decode()
        answer_dict = json.loads(cleaned)
        return answer_dict.get("md", "")
    except json.JSONDecodeError as e:
        raise ValueError(f"answer 字段不是合法 JSON：{e}")


# Init
mcp = FastMCP("i18n", stateless_http=True, host="0.0.0.0", port=8001)


@mcp.tool()
def extract_execl(file_url: str) -> str:
    """
    下载一个 Excel 文件并提取其表格内容，返回 Markdown 格式的字符串。

    参数:
        file_url: Excel 文件的直链 URL

    返回:
        表格内容，格式为 Markdown 表格字符串。
    """

    # 下载 Excel 文件
    response = requests.get(file_url)
    response.raise_for_status()

    # 使用临时文件保存内容
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_file:
        tmp_file.write(response.content)
        tmp_file.flush()

        # 读取 Excel 文件
        df = pd.read_excel(tmp_file.name, engine='openpyxl')

    # 排除所有项目都为空的行
    df = df.replace("", pd.NA).dropna(how="all")

    # 生成 Markdown 表格
    markdown = "|" + "|".join(df.columns) + "|\n"
    markdown += "|" + "|".join(["---"] * len(df.columns)) + "|\n"
    for _, row in df.iterrows():
        row_values = ["" if pd.isna(cell) else str(cell) for cell in row]
        markdown += "|" + "|".join(row_values) + "|\n"
    return markdown


@mcp.tool()
def process_excel(markdown_table: str) -> dict:
    """
    调用 Chatflow 服务翻译 Markdown 表格内容，并将结果生成新的 Excel 文件。

    参数:
        markdown_table: Markdown 表格字符串，需符合表格结构要求

    返回:
        包含翻译后 Excel 文件链接的字典对象，格式为:
        {
            "type": "link",
            "name": "<文件名称>",
            "url": "<文件链接>",
            "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
    """

    # 1. 调用 Chatflow
    processed_markdown = call_chatflow_with_markdown(markdown_table)

    # 2. 将返回的 Markdown 表格解析为 DataFrame
    lines = processed_markdown.strip().splitlines()
    if len(lines) < 3:
        raise ValueError("返回的 Markdown 表格无效")

    headers = [col.strip() for col in lines[0].split('|')]
    headers = headers[1:-1]
    rows = []
    for line in lines[2:]:
        parts = [cell.strip() for cell in line.split('|')]
        parts = parts[1:-1]
        rows.append(parts)

    df = pd.DataFrame(rows, columns=headers)

    # 创建 static 目录（如果不存在）
    os.makedirs("static", exist_ok=True)

    # 用唯一文件名保存
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
    file_name = f"translated_{timestamp}.xlsx"
    file_path = os.path.join("static", file_name)

    # 保存文件
    df.to_excel(file_path, index=False, engine="openpyxl")

    # 构造文件 URL（假设运行在 http://106.15.201.186:8001）
    file_url = f"http://106.15.201.186:8001/static/{file_name}"

    return {
        "type": "link",
        "name": file_name,
        "url": file_url,
        "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }


@mcp.tool()
def upload_and_process_excel(file_url: str) -> dict:
    """
    一站式处理 Excel 文件的工具：下载 Excel、提取为 Markdown、调用翻译服务并返回处理后的 Excel 文件。

    参数:
        file_url: Excel 文件的直链 URL

    返回:
        翻译后 Excel 文件链接的字典对象，格式为:
        {
            "type": "link",
            "name": "<文件名称>",
            "url": "<文件链接>",
            "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
    """

    extracted_data = extract_execl(file_url=file_url)
    processed_file = process_excel(markdown_table=extracted_data)
    return processed_file


if __name__ == "__main__":
    app = mcp.streamable_http_app()
    os.makedirs("static", exist_ok=True)
    app.mount("/static", StaticFiles(directory="static"), name="static")
    # Initialize and run the server
    mcp.run(transport='streamable-http')
