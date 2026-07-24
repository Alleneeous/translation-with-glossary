# 🌐 术语表辅助翻译工具

基于 DeepSeek API + Streamlit 的文档翻译 Web 应用。
上传文档和术语表 → AI 翻译时自动匹配并替换术语 → 下载术语一致的译文。

**术语表始终留在服务器端，用户只能看到翻译结果。**

---

## 本地运行

```bash
cd translate-app
pip install -r requirements.txt
# 设置 API Key
export DEEPSEEK_API_KEY="sk-..."
streamlit run app.py
```

浏览器访问 `http://localhost:8501`。

> API Key 通过环境变量 `DEEPSEEK_API_KEY` 传入，不会写死在代码里。本地运行时也可在侧边栏手动填入。

---

## 分发方式

### 方式一：Hugging Face Spaces（推荐 · 免费）

1. 访问 [huggingface.co/new-space](https://huggingface.co/new-space)
2. Space SDK 选 **Streamlit**，可见性建议选 **Private**
3. 将 `translate-app/` 目录下所有文件上传到 Space
4. 在 Space 设置 → **Repository Secrets** 中添加：
   ```
   DEEPSEEK_API_KEY = "sk-你的DeepSeek-API-Key"
   ```
5. 获得 URL：`https://huggingface.co/spaces/你的用户名/translate-app`
6. 发给用户的就是这个 URL

> 即使用 Public Space，API Key 通过 Secret 注入也不会暴露。

### 方式二：Streamlit Cloud

1. 将 `translate-app/` 推到 GitHub 仓库
2. 访问 [share.streamlit.io](https://share.streamlit.io)，用 GitHub 登录
3. New app → 选仓库 → 主文件 `app.py`
4. 在 App 设置 → **Secrets** 中添加：
   ```
   DEEPSEEK_API_KEY = "sk-你的DeepSeek-API-Key"
   ```
5. 获得 `https://your-app.streamlit.app`

### 方式三：内网部署

```bash
cd translate-app
pip install -r requirements.txt
export DEEPSEEK_API_KEY="sk-你的DeepSeek-API-Key"
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

同事通过 `http://<服务器IP>:8501` 访问。

### 方式四：Docker

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0"]
```

```bash
docker build -t translate-app .
docker run -p 8501:8501 -e DEEPSEEK_API_KEY="sk-..." translate-app
```

---

## 💰 费用参考

DeepSeek API 价格（比 Claude 便宜很多）：

| 模型 | 输入 | 输出 | 翻译100万字的费用 |
|------|------|------|-------------------|
| deepseek-v4-flash | ¥1 / 百万token | ¥2 / 百万token | 约 ¥5-10 |
| deepseek-v4-pro | ¥3 / 百万token | ¥6 / 百万token | 约 ¥15-30 |

---

## 文件结构

```
translate-app/
├── app.py               # Streamlit 主程序
├── glossary_helper.py   # 术语表处理逻辑（保护/恢复占位符）
├── glossary.xlsx        # (可选) 内置术语表，部署时放进去即可
├── requirements.txt     # Python 依赖
└── README.md            # 本文件
```
