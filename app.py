import os
import tempfile
import re
import base64
import requests
import hashlib
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
import zipfile
import traceback

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

def upload_image_to_github(image_data: bytes, image_name: str) -> str:
    """上传图片到 GitHub 仓库，返回 CDN 链接"""
    try:
        # 生成唯一文件名
        hash_name = hashlib.md5(image_data).hexdigest()[:12]
        safe_name = f"{hash_name}_{image_name}"
        
        content = base64.b64encode(image_data).decode("utf-8")
        
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/images/{safe_name}"
        headers = {
            "Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json"
        }
        
        # 先检查文件是否已存在
        check_resp = requests.get(url, headers=headers)
        if check_resp.status_code == 200:
            # 文件已存在，直接返回 CDN 链接
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
        else:
            raise Exception(f"GitHub API returned {resp.status_code}: {resp.text}")
            
    except Exception as e:
        print(f"Upload error: {str(e)}")
        raise

def extract_images_from_docx(docx_path: str) -> dict:
    """从 docx 文件中提取所有图片"""
    images = {}
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            for file_info in docx_zip.filelist:
                if file_info.filename.startswith('word/media/') and not file_info.filename.endswith('/'):
                    image_name = os.path.basename(file_info.filename)
                    image_data = docx_zip.read(file_info.filename)
                    images[image_name] = image_data
    except Exception as e:
        print(f"Error extracting images: {e}")
    
    return images

def docx_to_markdown(docx_path: str) -> str:
    """将 docx 转换为 Markdown"""
    try:
        doc = Document(docx_path)
        markdown_lines = []
        
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue
            
            # 检查是否是标题
            if paragraph.style and paragraph.style.name:
                style_name = paragraph.style.name.lower()
                if 'heading' in style_name:
                    # 提取标题级别
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
            
            # 普通段落
            markdown_lines.append(text)
        
        # 处理表格
        for table in doc.tables:
            if len(table.rows) > 0:
                markdown_lines.append("")  # 空行分隔
                
                # 表头
                header_cells = [cell.text.strip() for cell in table.rows[0].cells]
                markdown_lines.append("| " + " | ".join(header_cells) + " |")
                markdown_lines.append("|" + "|".join([" --- " for _ in header_cells]) + "|")
                
                # 数据行
                for row in table.rows[1:]:
                    cells = [cell.text.strip() for cell in row.cells]
                    markdown_lines.append("| " + " | ".join(cells) + " |")
                
                markdown_lines.append("")
        
        return "\n\n".join(markdown_lines) if markdown_lines else "# 转换结果\n\n文档内容为空"
        
    except Exception as e:
        print(f"Error converting docx: {e}")
        raise

@app.post("/convert")
async def convert_docx(file: UploadFile = File(...)):
    """转换 Word 文档为 Markdown"""
    
    # 详细的请求日志
    print(f"收到文件: {file.filename}, 大小: {file.size if file.size else 'unknown'}")
    
    # 检查文件类型
    if not file.filename.lower().endswith('.docx'):
        raise HTTPException(status_code=400, detail="只支持 .docx 格式的文件")
    
    # 限制文件大小 (10MB)
    if file.size and file.size > 10 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="文件大小不能超过 10MB")
    
    # 检查环境变量
    if not GITHUB_TOKEN:
        raise HTTPException(status_code=500, detail="服务器配置错误: 缺少 GitHub Token")
    if not GITHUB_REPO:
        raise HTTPException(status_code=500, detail="服务器配置错误: 缺少 GitHub 仓库信息")
    
    # 创建临时目录
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file.filename)
        
        # 保存上传的文件
        try:
            content = await file.read()
            print(f"文件读取成功，大小: {len(content)} 字节")
            
            with open(input_path, "wb") as f:
                f.write(content)
            print(f"文件保存成功: {input_path}")
            
        except Exception as e:
            print(f"文件保存失败: {e}")
            raise HTTPException(status_code=500, detail=f"文件保存失败: {str(e)}")
        
        markdown_content = ""
        
        try:
            # 步骤1: 转换为 Markdown
            print("开始转换 Word 到 Markdown...")
            markdown_content = docx_to_markdown(input_path)
            print(f"转换成功，内容长度: {len(markdown_content)}")
            
            # 步骤2: 提取图片
            print("开始提取图片...")
            images = extract_images_from_docx(input_path)
            print(f"找到 {len(images)} 张图片")
            
            # 步骤3: 上传图片
            if images:
                markdown_content += "\n\n## 图片附件\n\n"
                for img_name, img_data in images.items():
                    try:
                        print(f"上传图片: {img_name}")
                        cdn_url = upload_image_to_github(img_data, img_name)
                        markdown_content += f"![{img_name}]({cdn_url})\n\n"
                        print(f"图片上传成功: {cdn_url}")
                    except Exception as e:
                        error_msg = f"图片 {img_name} 上传失败: {str(e)}"
                        print(error_msg)
                        markdown_content += f"\n*{error_msg}*\n\n"
            
            # 步骤4: 确保有内容返回
            if not markdown_content.strip():
                markdown_content = "# 转换结果\n\n文档内容为空或无法解析"
            
            # 步骤5: 保存结果文件
            output_path = os.path.join(tmpdir, "converted.md")
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(markdown_content)
            
            print(f"Markdown 文件保存成功: {output_path}")
            
            # 步骤6: 返回文件
            return FileResponse(
                output_path,
                media_type="text/markdown",
                filename=Path(file.filename).stem + ".md"
            )
            
        except Exception as e:
            # 打印完整错误堆栈
            error_detail = traceback.format_exc()
            print(f"转换失败:\n{error_detail}")
            
            # 返回友好的错误信息
            raise HTTPException(
                status_code=500, 
                detail=f"转换失败: {str(e)}\n\n详细信息请查看服务器日志"
            )

@app.get("/")
async def root():
    return {
        "message": "Word to Markdown Converter 运行中",
        "version": "2.0",
        "status": "ok",
        "endpoints": {
            "convert": "POST /convert - 上传 .docx 文件",
            "health": "GET /health - 健康检查"
        }
    }

@app.get("/health")
async def health():
    """健康检查端点"""
    return {
        "status": "healthy",
        "github_repo": GITHUB_REPO if GITHUB_REPO else "未配置",
        "github_token_configured": bool(GITHUB_TOKEN)
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)