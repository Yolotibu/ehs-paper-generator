import streamlit as st
import requests
import re
import json
from datetime import datetime
import docx
import PyPDF2
from pptx import Presentation
import os

# 密钥本地保存文件
KEY_FILE = "api_key_cache.json"
def load_cached_key():
    if os.path.exists(KEY_FILE):
        try:
            with open(KEY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("api_key", "")
        except:
            return ""
    return ""

def save_cached_key(key):
    with open(KEY_FILE, "w", encoding="utf-8") as f:
        json.dump({"api_key": key}, f, ensure_ascii=False)

def clear_cached_key():
    if os.path.exists(KEY_FILE):
        os.remove(KEY_FILE)

# 读取单个文件，支持 txt / docx / pdf / pptx
def read_one_file(uploaded_file):
    text = ""
    fname = uploaded_file.name.lower()
    if fname.endswith(".txt"):
        text = uploaded_file.read().decode("utf-8", errors="ignore")
    elif fname.endswith(".docx"):
        doc = docx.Document(uploaded_file)
        text = "\n".join([para.text for para in doc.paragraphs])
    elif fname.endswith(".pdf"):
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
    elif fname.endswith(".pptx"):
        prs = Presentation(uploaded_file)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + "\n"
    return text

# 读取多个文件，合并全部文本
def read_multi_files(uploaded_files):
    all_text = ""
    for f in uploaded_files:
        all_text += f"\n===== 文件：{f.name} =====\n"
        all_text += read_one_file(f)
    return all_text

# 单个题型生成 + 校验
def gen_one_type(type_name, qty, diff, doc_text, headers, model_name):
    if type_name == "填空题":
        prompt = f"""
依据下面文档内容，只生成【{type_name}】，一共{qty}道，难度：{diff}。
⚠️ 重点：题目里只用____横线作为空白，**绝对不能把答案写在题目里面！**
不要别的题型、不要标题、不要分值、不要答案、不要解析。
只输出题目，题号从1开始连续编号。
难度定义：
简单：基础概念、定义；
中等：操作流程、管理要求；
困难：隐患判定、风险辨析、法规细节、事故处置。
文档：
{doc_text[:6000]}
"""
    else:
        prompt = f"""
依据下面文档内容，只生成【{type_name}】，一共{qty}道，难度：{diff}。
不要别的题型、不要标题、不要分值、不要答案、不要解析。
只输出题目，题号从1开始连续编号。
难度定义：
简单：基础概念、定义；
中等：操作流程、管理要求；
困难：隐患判定、风险辨析、法规细节、事故处置。
文档：
{doc_text[:6000]}
"""
    for _ in range(2):
        resp = requests.post(
            "https://api.siliconflow.cn/v1/chat/completions",
            headers=headers,
            json={
                "model": model_name,
                "messages": [{"role":"user", "content": prompt}],
                "temperature":0.05,
                "max_tokens":2000
            },
            timeout=90
        )
        resp.raise_for_status()
        res_txt = resp.json()["choices"][0]["message"]["content"]
        cnt = len(re.findall(r"\d+\.", res_txt))
        if cnt >= qty:
            return res_txt
    return res_txt

st.set_page_config(page_title="EHS整套试卷生成器", layout="wide")
st.title("📄 EHS培训整套试卷生成器")
st.info("💡 当前默认模型：Qwen/Qwen2.5-7B-Instruct（免费模型，不要随意切换付费模型）")

# 侧边栏
with st.sidebar:
    st.header("🔑 API配置")
    default_key = load_cached_key()
    api_key = st.text_input("硅基流动API Key", value=default_key, type="password")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("保存密钥"):
            save_cached_key(api_key)
            st.success("✅ 密钥已保存，下次自动加载！")
    with col2:
        if st.button("清空密钥"):
            clear_cached_key()
            st.rerun()

    model_name = st.text_input("模型名称", value="Qwen/Qwen2.5-7B-Instruct")
    st.divider()
    st.header("📝 试卷设置")
    pass_score = st.number_input("合格分数", min_value=40, max_value=99, value=60)
    st.markdown("> 试卷总分固定：100分")
    st.divider()
    st.header("📝 题型配置（勾选启用）")

    cfg = {}
    cfg["单选题"] = {
        "enable": st.checkbox("单选题", value=False),
        "num": st.number_input("单选数量", min_value=0, max_value=30, value=3),
        "diff": st.selectbox("单选难度", ["简单", "中等", "困难"], index=0)
    }
    cfg["多选题"] = {
        "enable": st.checkbox("多选题", value=True),
        "num": st.number_input("多选数量", min_value=0, max_value=30, value=2),
        "diff": st.selectbox("多选难度", ["简单", "中等", "困难"], index=1)
    }
    cfg["填空题"] = {
        "enable": st.checkbox("填空题", value=True),
        "num": st.number_input("填空数量", min_value=0, max_value=30, value=2),
        "diff": st.selectbox("填空难度", ["简单", "中等", "困难"], index=0)
    }
    cfg["判断题"] = {
        "enable": st.checkbox("判断题", value=True),
        "num": st.number_input("判断数量", min_value=0, max_value=30, value=3),
        "diff": st.selectbox("判断难度", ["简单", "中等", "困难"], index=0)
    }
    cfg["简答题"] = {
        "enable": st.checkbox("简答题", value=True),
        "num": st.number_input("简答数量", min_value=0, max_value=10, value=1),
        "diff": st.selectbox("简答难度", ["简单", "中等", "困难"], index=2)
    }

upload_files = st.file_uploader("📎 上传培训文档（txt / docx / pdf / pptx，支持多选）",
                               type=["txt","docx","pdf","pptx"],
                               accept_multiple_files=True)
doc_content = ""
if upload_files:
    doc_content = read_multi_files(upload_files)
    st.success("✅ 文件读取成功！")
    with st.expander("预览文档片段"):
        st.text(doc_content[:1500] + ("……" if len(doc_content) > 1500 else ""))

# 获取启用题型列表，顺序固定
type_order = ["单选题","多选题","填空题","判断题","简答题"]
active_list = []
for t in type_order:
    if cfg[t]["enable"] and cfg[t]["num"]>0:
        active_list.append((t, cfg[t]["num"], cfg[t]["diff"]))

if st.button("✨ 生成整套试卷"):
    if not api_key.strip():
        st.warning("请填写API Key！")
    elif not doc_content.strip():
        st.warning("请上传培训文档！")
    elif len(active_list) == 0:
        st.warning("至少勾选一种题型！")
    else:
        headers = {"Authorization":f"Bearer {api_key.strip()}", "Content-Type":"application/json"}
        st.info("正在逐个题型生成，会稍慢，请耐心等待……")
        parts = {}
        for tname, qty, diff in active_list:
            with st.spinner(f"正在生成：{tname} {qty}道"):
                parts[tname] = gen_one_type(tname, qty, diff, doc_content, headers, model_name)

        # 组装试卷正文，交给模型统一分配分值
        paper_raw = ""
        for tname, qty, diff in active_list:
            paper_raw += f"{tname}\n{parts[tname]}\n\n"

        prompt_format = f"""
下面是EHS试题，整理成标准试卷，总分固定100分，合格线：{pass_score}分。
【强制排版规则】
1. 试卷开头写标题：EHS培训考核试卷，总分：100分，合格分数线：{pass_score}分；
2. 大题序号必须连续递增：第一个题型是【一、XX题】，第二个题型【二、XX题】，第三个【三、XX题】，依次往后，不能全部都是“一、”；
3. 给每一道题合理分配分值，全部题目分值相加严格等于100分；
4. 只输出空白试卷，**不要答案、不要解析，填空题保持横线空白，严禁填入答案**；
5. 保留原有题目，不要修改、不要新增、不要删减题目。
试题内容：
{paper_raw}
"""
        with st.spinner("统一排版、分配分值……"):
            resp_fmt = requests.post(
                "https://api.siliconflow.cn/v1/chat/completions",
                headers=headers,
                json={
                    "model": model_name,
                    "messages": [{"role":"user", "content": prompt_format}],
                    "temperature":0.05,
                    "max_tokens":3000
                },
                timeout=90
            )
            resp_fmt.raise_for_status()
            final_paper = resp_fmt.json()["choices"][0]["message"]["content"]

        # 生成答案解析
        prompt_ans = f"""
为下面这份EHS试卷写出【参考答案+简短解析】，题号一一对应。
大题序号和试卷保持一致：一、二、三、四顺序对应。
只输出答案和解析，不要重新抄写题目。
试卷：
{final_paper}
"""
        with st.spinner("生成参考答案与解析……"):
            resp_ans = requests.post(
                "https://api.siliconflow.cn/v1/chat/completions",
                headers=headers,
                json={
                    "model": model_name,
                    "messages": [{"role":"user", "content": prompt_ans}],
                    "temperature":0.05,
                    "max_tokens":2800
                },
                timeout=90
            )
            resp_ans.raise_for_status()
            ans_text = resp_ans.json()["choices"][0]["message"]["content"]

        full_out = f"【试卷（无答案）】\n{final_paper}\n\n===== 参考答案+解析 =====\n{ans_text}"
        st.success("✅ 整套试卷生成完成！")
        st.text_area("试卷+答案（直接复制）", value=full_out, height=1000)
        save_name = f"EHS整套试卷_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(save_name, "w", encoding="utf-8") as f:
            f.write(full_out)
        st.info(f"📁 文件已保存到桌面：{save_name}")
