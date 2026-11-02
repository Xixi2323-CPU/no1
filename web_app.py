import streamlit as st
from openai import OpenAI
import io
import pandas as pd
from docx2python import docx2python
import json

st.set_page_config(page_title="AI 精准表单提取", page_icon="📑")
st.title("📑 复杂实验报告/表单提取器 (内测版)")
st.write("上传带合并单元格的 Word 表单，AI 自动提取成 Excel 数据！")

api_key = st.text_input("请输入你的 DeepSeek API Key", type="password")
uploaded_file = st.file_uploader("请上传 Word 文件", type=["docx"])

if st.button("开始精准提取"):
    if not api_key:
        st.warning("请先输入 API Key！")
    elif not uploaded_file:
        st.warning("请先上传文件！")
    else:
        with st.spinner("正在解析文档结构，请稍候..."):
            try:
                # 1. 用 docx2python 深度解析文档
                file_bytes = io.BytesIO(uploaded_file.getvalue())
                doc_result = docx2python(file_bytes)
                
                # 提取纯文本用于让 AI 理解上下文
                raw_text = str(doc_result.text)
                
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                
                # 2. 极简提示词，强制要求 JSON
                prompt = """
                你是一个专业的数据提取器。请从下面的文本中提取一份完整的实验报告信息。
                输出必须是一个 JSON 数组，包含且仅包含 **一行** 数据。
                字段必须严格为：["实验名称", "专业", "学号", "班级", "姓名", "实验时间", "实验地点", "实验目的", "实验内容", "实验过程", "实验结果", "心得体会"]。
                如果文本中缺失某部分，填“无”。不要输出除了 JSON 之外的任何字。
                """
                
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": prompt},
                        {"role": "user", "content": f"待提取文本：\n{raw_text[:6000]}"} 
                    ],
                    response_format={"type": "json_object"} # 强制 AI 输出 JSON，绝不废话
                )
                
                result_text = response.choices[0].message.content.strip()
                # 解析 JSON
                data = json.loads(result_text)
                
                # 如果 AI 返回的是单层字典，把它包进列表里
                if isinstance(data, dict):
                    if "data" in data: data = data["data"]
                    else: data = [data]
                
                df = pd.DataFrame(data)
                df = df.astype(str) # 防止学号变科学计数法
                
                st.success("提取完成！请检查下方数据：")
                st.dataframe(df)
                
                # 3. 导出 Excel
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                
                st.download_button(
                    label="⬇️ 下载 Excel 文件 (.xlsx)",
                    data=excel_buffer.getvalue(),
                    file_name="提取结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"提取失败，错误信息：{e}")
