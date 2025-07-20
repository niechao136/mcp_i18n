import io
import json

import pandas as pd
import requests
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
    answer_str = result.get("data", {}).get("outputs", {}).get("answer", "")
    if not answer_str:
        raise ValueError("响应中未找到 answer 字段")

    try:
        answer_json = json.loads(answer_str)
        return answer_json.get("md", "")
    except json.JSONDecodeError:
        raise ValueError("answer 字段不是有效的 JSON 字符串")


# Init
mcp = FastMCP("i18n", stateless_http=True, host="0.0.0.0", port=8001)


@mcp.tool()
async def extract_execl(file: bytes) -> str:
    """提取 Excel 文件内容并以 Markdown 表格格式返回

    Args:
        file: 上传的 Execl 文件
    """
    df = pd.read_excel(file, engine='openpyxl')
    markdown = "| " + " | ".join(df.columns) + " |\n"
    markdown += "| " + " | ".join(["---"] * len(df.columns)) + " |\n"
    for _, row in df.iterrows():
        markdown += "| " + " | ".join(str(cell) for cell in row) + " |\n"
    return markdown


@mcp.tool()
async def process_excel(markdown_table: str) -> bytes:
    """从 Markdown 表格中调用 Chatflow 翻译文件，获取处理后的 Markdown，并转为 Excel 文件

    Args:
        markdown_table: Markdown 表格字符串
    """

    # 1. 调用 Chatflow
    processed_markdown = call_chatflow_with_markdown(markdown_table)

    # 2. 将返回的 Markdown 表格解析为 DataFrame
    lines = processed_markdown.strip().splitlines()
    if len(lines) < 3:
        raise ValueError("返回的 Markdown 表格无效")

    headers = [col.strip() for col in lines[0].strip('|').split('|')]
    rows = []
    for line in lines[2:]:
        parts = [cell.strip() for cell in line.strip('|').split('|')]
        rows.append(parts)

    df = pd.DataFrame(rows, columns=headers)

    # 3. 写入 Excel 文件
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    return output.getvalue()


@mcp.tool()
def upload_and_process_excel(file: bytes) -> bytes:
    """
    主工具：自动调用 extract -> process
    上传 Excel 文件后自动提取内容并进行翻译，返回新文件

    Args:
        file: 上传的 Execl 文件
    """
    extracted_data = extract_execl(file=file)
    processed_file = process_excel(content=extracted_data)
    return processed_file


if __name__ == "__main__":
    # Initialize and run the server
    mcp.run(transport='streamable-http')

