
import io, os, re, json, base64, subprocess, tempfile
from pathlib import Path
import streamlit as st
from pptx import Presentation
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from openai import OpenAI

st.set_page_config(page_title="PPT教学设计生成器 V3.2", page_icon="📚", layout="wide")
APP_PASSWORD = os.getenv("APP_PASSWORD", "").strip()

VISION_PRESETS = {
    "阿里云百炼 / 通义千问": "https://dashscope.aliyuncs.com/compatible-mode/v1",
    "火山引擎 / 豆包": "https://ark.cn-beijing.volces.com/api/v3",
    "OpenAI": "https://api.openai.com/v1",
    "自定义 OpenAI 兼容接口": "",
}
WRITER_PRESETS = {
    "DeepSeek": "https://api.deepseek.com",
    "阿里云百炼 / 通义千问": "https://dashscope.aliyuncs.com/compatible-mode/v1",
    "火山引擎 / 豆包": "https://ark.cn-beijing.volces.com/api/v3",
    "自定义 OpenAI 兼容接口": "",
}
SECTIONS = ["课题","课型","教材分析","学情分析","教学目标","教学重难点","教学重点","教学难点",
            "教学方法","教学资源","教学过程","课堂小结","课后作业","作业设计","板书设计","教学反思"]

VISION_PROMPT = """你是一名初中社会学科PPT分析助手。请逐页识别文字、图片、地图、图表、时间轴、史料截图、
任务、问题、合作学习和课堂检测。只依据可见内容，不补充PPT没有呈现的信息。
按页码输出：页面主题、文字要点、视觉信息、教学功能、可进入教案的教师活动和学生活动。"""

WRITER_SYSTEM = """你是初中社会学科教学设计编辑器。必须以原教学设计为底稿，在原框架上修改：
保留原栏目和顺序；PPT决定内容和流程；删除与PPT不匹配的旧内容；保留匹配内容；
补入PPT真实存在但原教案遗漏的任务、史料、地图、图片、图表、时间轴、检测；
教师活动、学生活动、设计意图必须对应；不得虚构学生反应、课堂效果、教材页码或PPT未出现材料。
没有真实课堂反馈时，教学反思必须写成预设性内容。
只返回合法JSON：
{"title":"课题","sections":[{"heading":"原栏目名","content":"最终正文"}],
"alignment_check":{"removed":[],"kept":[],"added":[],"warnings":[]}}"""

def login_gate():
    if not APP_PASSWORD or st.session_state.get("auth_ok"): return
    st.title("📚 PPT教学设计生成器")
    pwd = st.text_input("访问密码", type="password")
    if st.button("登录"):
        if pwd == APP_PASSWORD:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("密码错误")
    st.stop()

def extract_ppt_text(data):
    prs = Presentation(io.BytesIO(data))
    out=[]
    for i,s in enumerate(prs.slides,1):
        texts=[]
        for sh in s.shapes:
            if hasattr(sh,"text") and getattr(sh,"text","").strip():
                texts.append(sh.text.strip())
        out.append((i,"\n".join(texts)))
    return out

def extract_template(data):
    doc=Document(io.BytesIO(data))
    raw=[]; heads=[]
    def check(t):
        c=re.sub(r"\s+","",t)
        for k in SECTIONS:
            if c==k or c.startswith(k+"（") or c.startswith(k+"("):
                if k not in heads: heads.append(k)
    for p in doc.paragraphs:
        t=p.text.strip()
        if t: raw.append(t); check(t)
    for ti,tb in enumerate(doc.tables,1):
        raw.append(f"[模板表格{ti}]")
        for row in tb.rows:
            vals=[c.text.strip() for c in row.cells]
            raw.append(" | ".join(vals))
            for v in vals:
                if v: check(v)
    whole="\n".join(raw)
    if len(heads)<3:
        heads=[k for k in SECTIONS if k in whole]
    return whole, heads

def ppt_to_images(data):
    with tempfile.TemporaryDirectory() as td:
        td=Path(td); ppt=td/"slides.pptx"; ppt.write_bytes(data)
        r=subprocess.run(["libreoffice","--headless","--convert-to","pdf","--outdir",str(td),str(ppt)],
                         capture_output=True,text=True,timeout=180)
        pdf=td/"slides.pdf"
        if r.returncode!=0 or not pdf.exists(): raise RuntimeError("PPT转PDF失败")
        r2=subprocess.run(["pdftoppm","-png","-r","120",str(pdf),str(td/"page")],
                          capture_output=True,text=True,timeout=180)
        if r2.returncode!=0: raise RuntimeError("PDF转图片失败")
        imgs=[p.read_bytes() for p in sorted(td.glob("page-*.png"))]
        if not imgs: raise RuntimeError("未生成PPT页面图片")
        return imgs

def test_api(key, base, model):
    c=OpenAI(api_key=key,base_url=base.rstrip("/"))
    r=c.chat.completions.create(model=model,messages=[{"role":"user","content":"请只回复：连接成功"}],max_tokens=20)
    return r.choices[0].message.content

def vision_call(key,base,model,images):
    c=OpenAI(api_key=key,base_url=base.rstrip("/"))
    reports=[]
    for start in range(0,len(images),4):
        batch=images[start:start+4]
        content=[{"type":"text","text":VISION_PROMPT+f"\n本批次对应PPT第{start+1}至{start+len(batch)}页。"}]
        for idx,img in enumerate(batch,start+1):
            b64=base64.b64encode(img).decode()
            content.append({"type":"text","text":f"PPT第{idx}页"})
            content.append({"type":"image_url","image_url":{"url":"data:image/png;base64,"+b64}})
        r=c.chat.completions.create(model=model,messages=[{"role":"user","content":content}])
        reports.append(r.choices[0].message.content)
    return "\n\n".join(reports)

def writer_call(key,base,model,prompt):
    c=OpenAI(api_key=key,base_url=base.rstrip("/"))
    try:
        r=c.chat.completions.create(model=model,messages=[
            {"role":"system","content":WRITER_SYSTEM},{"role":"user","content":prompt}],
            response_format={"type":"json_object"})
    except Exception:
        r=c.chat.completions.create(model=model,messages=[
            {"role":"system","content":WRITER_SYSTEM},{"role":"user","content":prompt+"\n只返回合法JSON。"}])
    return r.choices[0].message.content

def safe_json(s):
    s=s.strip()
    if s.startswith("```"):
        s=re.sub(r"^```(?:json)?\s*","",s); s=re.sub(r"\s*```$","",s)
    return json.loads(s)

def make_docx(orig,result,heads,subject,grade,periods):
    doc=Document(io.BytesIO(orig))
    body=doc._element.body; sect=body.sectPr
    for ch in list(body):
        if ch is not sect: body.remove(ch)
    p=doc.add_paragraph(); p.alignment=1
    r=p.add_run(result.get("title") or "教学设计"); r.bold=True; r.font.size=Pt(16)
    r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),"宋体")
    sm={str(x.get("heading","")).strip():str(x.get("content","")).strip() for x in result.get("sections",[])}
    for h in heads:
        doc.add_heading(h,level=1)
        for line in (sm.get(h) or "需教师补充").splitlines():
            if line.strip(): doc.add_paragraph(line.strip())
    doc.add_paragraph(f"学科：{subject}    年级：{grade}    课时：{periods}")
    buf=io.BytesIO(); doc.save(buf); return buf.getvalue()

login_gate()
st.title("📚 PPT教学设计生成器 V3.2")
st.caption("多模型视觉API · 多模型写作API · 原框架校正 · 删除不匹配内容")

with st.sidebar:
    st.header("① 视觉模型")
    vp=st.selectbox("视觉服务",list(VISION_PRESETS))
    vk=st.text_input("视觉 API Key",type="password")
    vb=st.text_input("视觉 API Base URL",value=VISION_PRESETS[vp])
    vm=st.text_input("视觉模型名称")
    if st.button("测试视觉API",use_container_width=True):
        try: st.success(test_api(vk,vb,vm))
        except Exception as e: st.error(str(e))

    st.divider(); st.header("② 教案生成模型")
    wp=st.selectbox("写作服务",list(WRITER_PRESETS))
    wk=st.text_input("写作 API Key",type="password")
    wb=st.text_input("写作 API Base URL",value=WRITER_PRESETS[wp])
    wm=st.text_input("写作模型名称")
    if st.button("测试写作API",use_container_width=True):
        try: st.success(test_api(wk,wb,wm))
        except Exception as e: st.error(str(e))

    st.divider(); st.header("③ 教学设置")
    subject=st.selectbox("学科",["历史","地理","道德与法治","社会"])
    grade=st.text_input("年级","八年级")
    textbook=st.text_input("教材版本","统编版")
    periods=st.text_input("课时","1课时")
    extra=st.text_area("特殊要求","")

c1,c2=st.columns(2)
with c1: ppt=st.file_uploader("④ 上传本课PPT",type=["pptx"])
with c2: docx=st.file_uploader("⑤ 上传原教学设计",type=["docx"])

st.info("示例：通义千问视觉 + DeepSeek写作；豆包视觉 + DeepSeek写作；或任意OpenAI兼容视觉/文本接口。")

if st.button("🚀 开始生成最终教学设计",type="primary",use_container_width=True):
    if not all([vk,vb,vm,wk,wb,wm,ppt,docx]):
        st.error("请完整填写两组API配置并上传两个文件。"); st.stop()
    try:
        prog=st.progress(0,text="读取文件…")
        pb=ppt.getvalue(); db=docx.getvalue()
        slides=extract_ppt_text(pb); template,heads=extract_template(db)
        ppt_text="\n\n".join(f"【PPT第{i}页】\n{t or '（依赖视觉识读）'}" for i,t in slides)
        prog.progress(20,text="PPT转图片…")
        imgs=ppt_to_images(pb)
        prog.progress(45,text="视觉识读…")
        vr=vision_call(vk,vb,vm,imgs)
        prompt=f"""【基本信息】\n学科：{subject}\n年级：{grade}\n教材版本：{textbook}\n课时：{periods}\n特殊要求：{extra or '无'}
【原教学设计】\n{template}
【必须保留栏目顺序】\n{chr(10).join(f'{i+1}. {h}' for i,h in enumerate(heads))}
【PPT文字】\n{ppt_text}
【PPT视觉报告】\n{vr}
请按原框架校正：删除与PPT不匹配内容，保留匹配内容，补入PPT真实存在内容，最终可直接使用。"""
        prog.progress(70,text="校正教案…")
        result=safe_json(writer_call(wk,wb,wm,prompt))
        prog.progress(90,text="生成Word…")
        out=make_docx(db,result,heads,subject,grade,periods)
        prog.progress(100,text="完成")
        st.success("生成完成")
        sm={str(x.get("heading","")).strip():str(x.get("content","")).strip() for x in result.get("sections",[])}
        tabs=st.tabs(["📄 最终教学设计","🔍 修改核对","🖼️ PPT视觉报告"])
        with tabs[0]:
            st.header(result.get("title") or "教学设计")
            for h in heads: st.subheader(h); st.markdown(sm.get(h) or "需教师补充")
        with tabs[1]:
            c=result.get("alignment_check",{})
            for title,key in [("已删除","removed"),("已保留","kept"),("已补入","added"),("需确认","warnings")]:
                st.subheader(title)
                for x in c.get(key,[]) or ["无"]: st.write("- "+x)
        with tabs[2]: st.markdown(vr)
        st.download_button("⬇️ 下载最终Word教学设计",out,
            file_name=f"{result.get('title') or '教学设计'}_PPT适配版.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
    except Exception as e:
        st.error("生成失败"); st.exception(e)
