import streamlit as st
from openai import OpenAI
import io
import json
import pandas as pd
from docx2python import docx2python

st.set_page_config(page_title="AI 表单精准提取", page_icon="📑")
st.title("📑 复杂实验报告/表单提取器")
st.write("上传这种带合并单元格的 Word 表单，AI 会自动把整页内容精准提取成 Excel 的一行数据！")

api_key = st.text_input("请输入你的 DeepSeek API Key", type="password")
uploaded_file = st.file_uploader("请上传 Word 文件", type=["docx"])

if st.button("开始精准提取"):
    if not api_key:
        st.warning("请先输入 API Key！")
    elif not uploaded_file:
        st.warning("请先上传文件！")
    else:
        with st.spinner("正在深度解析文档结构，并呼叫 AI 填表..."):
            try:
                # 直接用 docx2python 提取所有纯文本，不搞复杂的遍历，避免报错！
                doc_content = docx2python(io.BytesIO(uploaded_file.getvalue()))
                
                # doc_content.text 会直接返回文档里所有的文字（包括表格里的），用换行符隔开
                raw_text = str(doc_content.text)
                
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                
                # 告诉 AI，我需要一行包含所有信息的 JSON
                prompt = """
                你是一个极度严谨的数据提取器。下面是一份从Word提取出的表单文本，由于排版原因可能有些杂乱。
                请仔细阅读，提取出以下所有字段：实验名称、专业、学号、班级、姓名、实验时间、实验地点、实验目的、实验内容、实验过程、实验结果、心得体会。
                
                要求：
                1. 输出一个 JSON 数组，里面只有 **一个** 字典对象（因为这是一份报告，对应Excel的一行）。
                2. 如果某项内容在文本中没有提到，请填“无”。
                3. 只输出纯 JSON，不要有任何 markdown 格式（不要 ```json）。
                """
                
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": prompt},
                        {"role": "user", "content": f"表单内容如下：\n{raw_text[:5000]}"} 
                    ]
                )
                
                result_text = response.choices[0].message.content.strip()
                # 清理可能存在的 markdown 标记
                if result_text.startswith("```json"):
                    result_text = result_text.replace("```json", "").replace("```", "").strip()
                
                # 解析 JSON
                data = json.loads(result_text)
                df = pd.DataFrame(data)
                df = df.astype(str) 
                
                st.success("提取成功！请检查：")
                st.dataframe(df)
                
                # 导出 Excel
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                
                st.download_button(
                    label="⬇️ 点击下载 Excel 文件 (.xlsx)",
                    data=excel_buffer.getvalue(),
                    file_name="AI提取结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except json.JSONDecodeError:
                st.error("AI 返回的格式不够标准，请重试。")
            except Exception as e:
                st.error(f"出错了：{e}")