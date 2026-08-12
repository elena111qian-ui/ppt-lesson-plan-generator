# PPT教学设计生成器 V3.0 — 云端网页版

这是一个无需在工作电脑安装 Python 的网页应用。

工作电脑只需要：
1. 打开浏览器；
2. 输入网页访问密码；
3. 上传 PPTX；
4. 上传原 DOCX 教学设计；
5. 点击生成；
6. 下载最终 Word。

所有 Python、LibreOffice、AI API 调用都运行在云端服务器。

---

## 核心流程

PPTX
→ 云端 LibreOffice 转 PDF
→ OpenAI 视觉模型整页识读
→ 获取文字 / 图片 / 地图 / 图表 / 史料 / 时间轴 / 教学任务

原教学设计 DOCX
→ 读取原栏目与顺序
→ 作为底稿

两者合并
→ DeepSeek 或 OpenAI 校正
→ 删除所有与 PPT 不匹配内容
→ 保留仍然匹配的内容
→ 补入 PPT 中真实存在但原教案遗漏的教学活动
→ 生成最终 Word

---

## 与 V2.2 的主要区别

- 不需要公司电脑安装 Python
- 不需要公司电脑安装 LibreOffice
- 不需要每次在浏览器输入 API Key
- API Key 存在服务器环境变量中
- 可设置网页访问密码
- Docker 中自带 LibreOffice 和中文字体
- 可以部署到 Render 或其他支持 Docker 的云平台

---

# 推荐部署：Render + GitHub

## 第一步：创建一个 GitHub 仓库

把本文件夹中的这些文件上传到 GitHub：

- app.py
- requirements.txt
- Dockerfile
- render.yaml
- .gitignore

不要上传 `.env`，也不要把真实 API Key 写进代码。

## 第二步：在 Render 创建 Web Service

1. 登录 Render。
2. New → Web Service。
3. 连接你的 GitHub 仓库。
4. 选择 Docker。
5. 创建服务。

也可以使用仓库中的 `render.yaml` 作为 Render Blueprint。

## 第三步：配置 Secrets / Environment Variables

在 Render 后台添加：

- `OPENAI_API_KEY`：用于视觉识读
- `DEEPSEEK_API_KEY`：用于 DeepSeek 教案生成（如使用DeepSeek）
- `APP_PASSWORD`：你自己设置的网页登录密码

可选：

- `VISION_MODEL=gpt-5.6`
- `DEEPSEEK_MODEL=deepseek-v4-flash`
- `OPENAI_WRITER_MODEL=gpt-5.6`

保存后重新部署。

## 第四步：日常使用

部署成功后 Render 会提供一个 HTTPS 网页地址。

以后在公司电脑：

打开网址
→ 输入 APP_PASSWORD
→ 上传 PPT
→ 上传原教学设计
→ 点击“开始生成最终教学设计”
→ 下载 Word

电脑无需安装 Python。

---

# 数据与安全

这个版本：

- 不内置数据库；
- 不主动把上传的 PPT / DOCX 保存到永久磁盘；
- 文件在一次生成请求中临时处理；
- API Key 只从服务器环境变量读取，不发送到浏览器页面；
- 建议设置 `APP_PASSWORD`，避免生成器成为公开网页。

但请注意：
PPT 和教学设计内容会发送给你所配置的 AI API 服务商进行处理。
如果学校或公司对教学材料、学生信息或内部资料有数据合规要求，应先确认允许使用外部云服务。

不要上传包含学生身份证号、家庭地址、成绩隐私等不必要的敏感个人信息。

---

# 为什么要用 Docker

云端需要 LibreOffice 把 PPTX 转成 PDF，才能尽可能保留地图、图表、SmartArt、图片和页面版式。
Docker 允许服务器在构建时安装 LibreOffice 和中文字体，因此工作电脑不需要安装任何东西。

---

# 其他部署平台

本项目是标准 Docker Web 应用，因此也可以部署到其他支持 Docker 的云服务器/平台。
核心要求是：

- 能构建 Dockerfile
- 能设置环境变量/Secrets
- 给应用提供一个 `PORT`
- 能访问 OpenAI / DeepSeek API

---

# 当前限制

1. Word 输出会尽量继承原 DOCX 的页面设置、页眉页脚和样式资源，但复杂表格的像素级原位替换仍不是完全复刻。
2. PPT 转 PDF 后视觉识读会增加模型输入量和 API 成本。
3. 超大 PPT 建议压缩图片后再上传。
4. 公司网络如果屏蔽部署平台域名，仍需由公司网络管理员放行；不建议绕过公司安全策略。
