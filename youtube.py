import yt_dlp
from comtypes.client import CreateObject
from pdf2docx import Converter
import sys

# 定义回调函数，用于更新和显示进度条
def progress_callback(current, total):
    progress = int(100 * current / total)
    sys.stdout.write(f"\r转换进度: {progress}%")
    sys.stdout.flush()

print("1:下载视频\t2:Word转PDF\n3:PDF转Word")
print("=============================================")
number = input("请输入序号：")

if number == '1':
    url = input("请输入网址：")
    ydl_opts = {
        'outtmpl': 'videos/%(title)s.%(ext)s',  # 下载的视频将被存储到videos文件夹，文件名为视频标题
    }

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        ydl.download([url])

    print("视频下载成功!已保存在videos文件夹中")

elif number == '2':
    # 注意：这个功能仅在Windows系统上且已安装Microsoft Word的情况下工作
    docx_path = input("请输入Word文件路径：")
    pdf_path = input("请输入PDF文件完整保存路径（包括文件名）：")

    # 确保pdf_path以.pdf结尾
    if not pdf_path.lower().endswith('.pdf'):
        pdf_path += '.pdf'

    word = CreateObject('Word.Application')
    word.Visible = False  # 运行Word的背景模式，不显示界面
    try:
        doc = word.Documents.Open(docx_path)
        # 获取文档页数，用于进度条显示
        total_pages = doc.ComputeStatistics(2)
        current_page = 0
        # 逐页保存为PDF
        for i in range(total_pages):
            doc.ExportAsFixedFormat(pdf_path, 17, Start=i + 1, End=i + 1)
            current_page += 1
            # 更新进度条
            progress_callback(current_page, total_pages)
        doc.Close()
        print(f"\nWord文档 {docx_path} 已成功转换为PDF并保存在 {pdf_path}")
    except Exception as e:
        print(f"转换Word文档时出错: {e}")
    finally:
        word.Quit()

elif number == '3':
    pdf_path = input("请输入PDF文件路径：")
    docx_path = input("请输入要保存的Word文件路径（包括文件名）：")

    # 创建一个转换器实例，并传入回调函数
    cv = Converter(pdf_path, progress_callback=progress_callback)

    # 转换全部页面
    cv.convert(docx_path, start=0, end=None)

    # 关闭转换器
    cv.close()

    print(f"\nPDF文档 {pdf_path} 已成功转换为Word文档并保存在 {docx_path}")
