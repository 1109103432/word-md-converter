import os
import subprocess
import win32com.client as win32
from pathlib import Path

def word_to_md_ultimate(input_docx, output_md, media_dir):
    """
    终极混合逻辑：Win32com 洗大纲 + Pandoc 转格式提图片
    """
    input_docx = os.path.abspath(input_docx)
    output_md = os.path.abspath(output_md)
    temp_docx = input_docx.replace(".docx", "_temp_fixed.docx")

    print(f"  -> 1. 正在调用 Word 原生引擎清洗大纲级别...")
    word = win32.DispatchEx('Word.Application')
    word.Visible = False
    
    try:
        doc = word.Documents.Open(input_docx)
        for p in doc.Paragraphs:
            lvl = p.OutlineLevel
            if 1 <= lvl <= 6:
                try:
                    p.Style = -1 - lvl
                except Exception:
                    pass 
        
        # 另存为临时文件
        doc.SaveAs2(temp_docx, 16)
        doc.Close(False)
    except Exception as e:
        print(f"  ❌ Word 引擎处理出错: {e}")
        return False
    finally:
        word.Quit()

    print(f"  -> 2. 正在调用 Pandoc 生成 Markdown 并提取图片...")
    try:
        # 调用 Pandoc 的核心命令
        # -t gfm: 保证表格完美兼容
        # --extract-media: 自动提取图片并更新 md 中的链接
        # --wrap=none: 防止中文段落被异常换行打断
        cmd = [
            'pandoc', temp_docx,
            '-f', 'docx',
            '-t', 'gfm',
            '-o', output_md,
            f'--extract-media={media_dir}',
            '--wrap=none'
        ]
        
        # 执行命令并捕获输出
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(f"  ✅ 转换成功！图片已存放至: {media_dir} 文件夹")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"  ❌ Pandoc 转换出错:\n{e.stderr}")
        return False
    except FileNotFoundError:
        print(f"  ❌ 找不到 Pandoc 程序，请确保已安装并配置了系统环境变量。")
        return False
    finally:
        # 清理临时文件
        if os.path.exists(temp_docx):
            os.remove(temp_docx)


def batch_convert_current_dir():
    current_dir = Path.cwd()
    print(f"🔍 正在扫描目录: {current_dir}")
    
    all_docx_files = list(current_dir.glob("*.docx"))
    
    valid_files = [
        f for f in all_docx_files 
        if not f.name.startswith("~$") and not f.name.endswith("_temp_fixed.docx")
    ]
    
    if not valid_files:
        print("⚠️ 当前目录下没有找到可转换的 .docx 文件！")
        return

    print(f"📋 共找到 {len(valid_files)} 个 Word 文档，准备开始转换...\n")
    
    success_count = 0
    for idx, file_path in enumerate(valid_files, 1):
        input_path = str(file_path)
        output_path = str(file_path.with_suffix('.md'))
        
        # 为了防止不同文档的图片重名混淆，给每个文档单独建一个同名的图片存储文件夹
        media_folder = f"{file_path.stem}_assets"
        
        print(f"[{idx}/{len(valid_files)}] 开始处理: {file_path.name}")
        
        if word_to_md_ultimate(input_path, output_path, media_folder):
            success_count += 1
            
    print(f"\n🎉 全部任务结束！成功转换 {success_count}/{len(valid_files)} 个文件。")

if __name__ == "__main__":
    batch_convert_current_dir()
    input("\n按回车键退出...")
