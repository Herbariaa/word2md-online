from fastapi.staticfiles import StaticFiles
import os
import tempfile
import subprocess
import re
import base64
import requests
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse

app = FastAPI(title="Word to Markdown Converter", version="1.0")
# 挂载静态文件目录（提供 index.html）
app.mount("/", StaticFiles(directory="static", html=True), name="static")
# 从环境变量读取配置
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPO = os.getenv("GITHUB_REPO")  # 格式: username/word2md-images

def upload_image_to_github(image_path: str, image_name: str) -> str:
    """上传图片到 GitHub 仓库，并返回 jsDelivr CDN 链接"""
    with open(image_path, "rb") as f:
        content = base64.b64encode(f.read()).decode("utf-8")

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/images/{image_name}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }
    data = {
        "message": f"Upload {image_name} via word2md",
        "content": content,
        "branch": "main"
    }

    resp = requests.put(url, json=data, headers=headers)
    if resp.status_code == 201:
        username, repo = GITHUB_REPO.split("/")
        return f"https://cdn.jsdelivr.net/gh/{username}/{repo}@main/images/{image_name}"
    else:
        raise Exception(f"Upload failed: {resp.text}")

@app.post("/convert")
async def convert_docx(file: UploadFile = File(...)):
    # 1. 限制文件类型和大小
    if not file.filename.lower().endswith(('.docx', '.doc')):
        raise HTTPException(status_code=400, detail="Only .docx or .doc files allowed")
    if file.size > 10 * 1024 * 1024:  # 10MB
        raise HTTPException(status_code=400, detail="File too large (max 10MB)")

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file.filename)
        with open(input_path, "wb") as f:
            f.write(await file.read())

        # 2. .docx → .html (LibreOffice)
        try:
            subprocess.run([
                "libreoffice", "--headless", "--convert-to", "html",
                "--outdir", tmpdir, input_path
            ], check=True, timeout=60)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")

        # 3. .html → .md (Pandoc)
        base_name = Path(file.filename).stem
        html_path = os.path.join(tmpdir, base_name + ".html")
        md_path = os.path.join(tmpdir, "output.md")
        subprocess.run(["pandoc", html_path, "-f", "html", "-t", "markdown", "-o", md_path], check=True)

        # 4. 处理图片
        image_dir = os.path.join(tmpdir, base_name + "_files")
        with open(md_path, "r", encoding="utf-8") as f:
            content = f.read()

        def replace_image(match):
            local_rel_path = match.group(1)
            img_filename = Path(local_rel_path).name
            local_img_path = os.path.join(image_dir, img_filename)
            if os.path.exists(local_img_path):
                try:
                    cdn_url = upload_image_to_github(local_img_path, img_filename)
                    return f"![]({cdn_url})"
                except Exception as e:
                    print(f"Upload error: {e}")
            return match.group(0)

        new_content = re.sub(r'!\$$[^)]*\$$(\S+)', replace_image, content)

        # 5. 返回结果
        final_md = os.path.join(tmpdir, "final_output.md")
        with open(final_md, "w", encoding="utf-8") as f:
            f.write(new_content)

        return FileResponse(final_md, filename="converted.md")