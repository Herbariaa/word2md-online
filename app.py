import os
import tempfile
import re
import base64
import requests
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml
import zipfile
from io import BytesIO

app = FastAPI(title="Word to Markdown Converter", version="2.0")

# 允许跨域（方便前端调用）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 从环境变量读取配置
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPO = os.getenv("GITHUB_REPO")  # 格式: username/word2md-images

def upload_image_to_github(image_data: bytes, image_name: str) -> str:
    """上传图片到 GitHub 仓库，返回 CDN 链接"""
    content = base64.b64encode(image_data).decode("utf-8")
    
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
    if resp.status_code in [200, 201]:
        username, repo = GITHUB_REPO.split("/")
        return f"https://cdn.jsdelivr.net/gh/{username}/{repo}@main/images/{image_name}"
    else:
        raise Exception(f"Upload failed: {resp.status_code} - {resp.text}")

def extract_images_from_docx(docx_path: str) -> dict:
    """从 docx 文件中提取所有图片"""
    images = {}
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # 查找所有图片文件
        for file_info in docx_zip.filelist:
            if file_info.filename.startswith('word/media/'):
                image_name = os.path.basename(file_info.filename)
                image_data = docx_zip.read(file_info.filename)
                images[image_name] = image_data
    return images

def docx_to_markdown(docx_path: str, images: dict) -> str:
    """将 docx 转换为 Markdown"""
    doc = Document(docx_path)
    markdown_lines = []
    
    # 图片计数器
    image_counter = 0
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        
        # 检查段落样式
        style_name = paragraph.style.name.lower() if paragraph.style else ""
        
        # 处理标题
        if style_name.startswith('heading'):
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
            else:
                level = 1
            
            if text:
                markdown_lines.append(f"{'#' * level} {text}")
                markdown_lines.append("")
        
        # 处理列表
        elif style_name.startswith('list'):
            if text:
                markdown_lines.append(f"- {text}")
        
        # 处理普通段落
        else:
            if text:
                # 处理粗体和斜体
                # 注意：python-docx 的粗体/斜体处理比较复杂
                # 这里简化处理
                markdown_lines.append(text)
                markdown_lines.append("")
    
    # 处理表格
    for table in doc.tables:
        markdown_lines.append("")  # 空行分隔
        
        # 处理表头（第一行）
        if len(table.rows) > 0:
            header_cells = [cell.text.strip() for cell in table.rows[0].cells]
            markdown_lines.append("| " + " | ".join(header_cells) + " |")
            markdown_lines.append("|" + "|".join([" --- " for _ in header_cells]) + "|")
            
            # 处理数据行
            for row in table.rows[1:]:
                cells = [cell.text.strip() for cell in row.cells]
                markdown_lines.append("| " + " | ".join(cells) + " |")
        
        markdown_lines.append("")
    
    return "\n".join(markdown_lines)

def process_inline_images(docx_path: str, markdown_content: str) -> str:
    """处理文档中的内联图片"""
    doc = Document(docx_path)
    images = extract_images_from_docx(docx_path)
    
    # 图片计数器
    img_index = 1
    image_replacements = []
    
    # 遍历所有段落查找图片
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            # 检查是否有图片
            if run._element.xpath('.//a:blip'):
                # 找到图片引用
                for blip in run._element.xpath('.//a:blip'):
                    embed_id = blip.get(qn('r:embed'))
                    if embed_id:
                        # 在 images 中查找对应的图片
                        for img_name, img_data in images.items():
                            try:
                                cdn_url = upload_image_to_github(img_data, f"img_{img_index}_{img_name}")
                                image_replacements.append((img_index, cdn_url))
                                img_index += 1
                                break
                            except Exception as e:
                                print(f"Upload error: {e}")
    
    # 如果找到了图片，在 Markdown 中插入图片引用
    if image_replacements:
        # 简单地在文档末尾添加图片
        markdown_content += "\n\n## 图片\n\n"
        for img_id, img_url in image_replacements:
            markdown_content += f"![图片{img_id}]({img_url})\n\n"
    
    return markdown_content

@app.post("/convert")
async def convert_docx(file: UploadFile = File(...)):
    """转换 Word 文档为 Markdown"""
    
    # 检查文件类型
    if not file.filename.lower().endswith(('.docx')):
        raise HTTPException(status_code=400, detail="Only .docx files allowed")
    
    # 限制文件大小 (10MB)
    if file.size > 10 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="File too large (max 10MB)")
    
    # 检查环境变量
    if not GITHUB_TOKEN or not GITHUB_REPO:
        raise HTTPException(status_code=500, detail="Server configuration error: missing GitHub credentials")
    
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file.filename)
        
        # 保存上传的文件
        with open(input_path, "wb") as f:
            f.write(await file.read())
        
        try:
            # 提取图片
            images = extract_images_from_docx(input_path)
            
            # 转换为 Markdown
            markdown_content = docx_to_markdown(input_path, images)
            
            # 处理并上传图片
            final_markdown = process_inline_images(input_path, markdown_content)
            
            # 如果没有任何图片上传成功，至少返回文本内容
            if not final_markdown.strip():
                final_markdown = "# 转换结果\n\n文档内容为空或无法解析"
            
            # 保存结果
            output_path = os.path.join(tmpdir, "converted.md")
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(final_markdown)
            
            return FileResponse(
                output_path,
                media_type="text/markdown",
                filename=Path(file.filename).stem + ".md"
            )
            
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")

@app.get("/")
async def root():
    return {
        "message": "Word to Markdown Converter is running",
        "version": "2.0",
        "status": "ok",
        "endpoints": {
            "convert": "POST /convert - Upload .docx file",
            "health": "GET /health"
        }
    }

@app.get("/health")
async def health():
    """健康检查端点"""
    return {
        "status": "healthy",
        "github_repo": GITHUB_REPO if GITHUB_REPO else "not configured"
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)