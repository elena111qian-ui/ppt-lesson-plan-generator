# V3.2 多模型视觉API版

视觉识读支持：
- 阿里云百炼 / 通义千问
- 火山引擎 / 豆包
- OpenAI
- 自定义 OpenAI 兼容接口

教案生成支持：
- DeepSeek
- 阿里云百炼 / 通义千问
- 火山引擎 / 豆包
- 自定义 OpenAI 兼容接口

如果你已经部署 V3.1 到 Render，只需要替换 GitHub 仓库中的：
1. app.py
2. Dockerfile
3. requirements.txt

然后 Commit changes。Render 开启 Auto Deploy 时会自动重新部署。

注意：V3.2 新增 poppler-utils，因此必须替换 Dockerfile。
