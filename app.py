import os
import tempfile
import base64
import hashlib
import zipfile
import traceback
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
import requests

app = FastAPI(title="Word to Markdown Converter", version="2.0")

# 允许跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 从环境变量读取配置
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPO = os.getenv("GITHUB_REPO")

# ========== 辅助函数 ==========
def upload_image_to_github(image_data: bytes, image_name: str) -> str:
    try:
        hash_name = hashlib.md5(image_data).hexdigest()[:12]
        safe_name = f"{hash_name}_{image_name}"
        content = base64.b64encode(image_data).decode("utf-8")
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/images/{safe_name}"
        headers = {
            "Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json"
        }
        
        # 检查文件是否已存在
        check_resp = requests.get(url, headers=headers)
        if check_resp.status_code == 200:
            username, repo = GITHUB_REPO.split("/")
            return f"https://cdn.jsdelivr.net/gh/{username}/{repo}@main/images/{safe_name}"
        
        data = {
            "message": f"Upload {safe_name} via word2md",
            "content": content,
            "branch": "main"
        }
        resp = requests.put(url, json=data, headers=headers)
        
        if resp.status_code in [200, 201]:
            username, repo = GITHUB_REPO.split("/")
            return f"https://cdn.jsdelivr.net/gh/{username}/{repo}@main/images/{safe_name}"
        raise Exception(f"GitHub API returned {resp.status_code}")
    except Exception as e:
        print(f"Upload error: {str(e)}")
        raise

def extract_images_from_docx(docx_path: str) -> dict:
    images = {}
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            for file_info in docx_zip.filelist:
                if file_info.filename.startswith('word/media/') and not file_info.filename.endswith('/'):
                    image_name = os.path.basename(file_info.filename)
                    images[image_name] = docx_zip.read(file_info.filename)
    except Exception as e:
        print(f"Error extracting images: {e}")
    return images

def docx_to_markdown(docx_path: str) -> str:
    try:
        doc = Document(docx_path)
        markdown_lines = []
        
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue
            
            if paragraph.style and paragraph.style.name:
                style_name = paragraph.style.name.lower()
                if 'heading' in style_name:
                    level = 1
                    if '1' in style_name:
                        level = 1
                    elif '2' in style_name:
                        level = 2
                    elif '3' in style_name:
                        level = 3
                    elif '4' in style_name:
                        level = 4
                    elif '5' in style_name:
                        level = 5
                    markdown_lines.append(f"{'#' * level} {text}")
                    continue
            
            markdown_lines.append(text)
        
        # 处理表格
        for table in doc.tables:
            if len(table.rows) > 0:
                markdown_lines.append("")
                header_cells = [cell.text.strip() for cell in table.rows[0].cells]
                markdown_lines.append("| " + " | ".join(header_cells) + " |")
                markdown_lines.append("|" + "|".join([" --- " for _ in header_cells]) + "|")
                for row in table.rows[1:]:
                    cells = [cell.text.strip() for cell in row.cells]
                    markdown_lines.append("| " + " | ".join(cells) + " |")
                markdown_lines.append("")
        
        return "\n\n".join(markdown_lines) if markdown_lines else "# 转换结果\n\n文档内容为空"
    except Exception as e:
        print(f"Error converting docx: {e}")
        raise

# ========== API 端点 ==========
@app.get("/")
async def root():
    """返回前端页面"""
    possible_paths = [
        Path(__file__).parent / "static" / "index.html",
        Path("/app/static/index.html"),
        Path("static/index.html"),
    ]
    
    for path in possible_paths:
        if path.exists():
            print(f"找到静态文件: {path}")
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                return HTMLResponse(content=content)
            except Exception as e:
                print(f"读取文件失败: {e}")
                continue
    
    # 如果没有找到静态文件，返回简单的 HTML
    return HTMLResponse(content="""
    <html>
        <body style="font-family: sans-serif; text-align: center; padding: 50px;">
            <h1>📄 Word to Markdown Converter</h1>
            <p>服务运行正常，但前端页面未找到。</p>
            <p>请使用 POST /convert 接口上传文件。</p>
            <p><a href="/health">健康检查</a></p>
        </body>
    </html>
    """)

@app.post("/convert")
async def convert_docx(file: UploadFile = File(...)):
    """转换 Word 文档为 Markdown"""
    print(f"收到文件: {file.filename}")
    
    if not file.filename.lower().endswith('.docx'):
        raise HTTPException(status_code=400, detail="只支持 .docx 格式的文件")
    
    if file.size and file.size > 10 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="文件大小不能超过 10MB")
    
    if not GITHUB_TOKEN or not GITHUB_REPO:
        raise HTTPException(status_code=500, detail="服务器配置错误")
    
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file.filename)
        
        try:
            content = await file.read()
            with open(input_path, "wb") as f:
                f.write(content)
            print(f"文件保存成功: {input_path}, 大小: {len(content)} 字节")
        except Exception as e:
            print(f"文件保存失败: {e}")
            raise HTTPException(status_code=500, detail=f"文件保存失败: {str(e)}")
        
        try:
            # 1. 转换为 Markdown
            markdown_content = docx_to_markdown(input_path)
            print(f"转换成功，内容长度: {len(markdown_content)}")
            
            # 2. 提取图片
            images = extract_images_from_docx(input_path)
            print(f"找到 {len(images)} 张图片")
            
            # 3. 上传图片并添加到 Markdown
            if images:
                markdown_content += "\n\n## 📷 图片附件\n\n"
                for idx, (img_name, img_data) in enumerate(images.items(), 1):
                    try:
                        cdn_url = upload_image_to_github(img_data, img_name)
                        markdown_content += f"![图片{idx}]({cdn_url})\n\n"
                        print(f"图片上传成功: {img_name} -> {cdn_url}")
                    except Exception as e:
                        error_msg = f"图片 {img_name} 上传失败: {str(e)}"
                        print(error_msg)
                        markdown_content += f"\n*⚠️ {error_msg}*\n\n"
            
            # 4. 确保有内容
            if not markdown_content.strip():
                markdown_content = "# 转换结果\n\n文档内容为空或无法解析"
            
            # 5. 保存结果文件（注意：使用一致的文件名）
            output_path = os.path.join(tmpdir, "converted.md")
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(markdown_content)
            
            print(f"Markdown 文件保存成功: {output_path}")
            print(f"文件大小: {os.path.getsize(output_path)} 字节")
            
            # 6. 返回文件
            return FileResponse(
                output_path,
                media_type="text/markdown",
                filename=Path(file.filename).stem + ".md"
            )
            
        except Exception as e:
            error_detail = traceback.format_exc()
            print(f"转换失败:\n{error_detail}")
            raise HTTPException(status_code=500, detail=f"转换失败: {str(e)}")

@app.get("/health")
async def health():
    """健康检查端点"""
    return {
        "status": "healthy",
        "github_repo": GITHUB_REPO,
        "github_token_configured": bool(GITHUB_TOKEN),
        "version": "2.0"
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)