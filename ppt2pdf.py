'''
ppt,pps,pptx files to pdf
many ebooks is ppt file, calibre can't manage,
so this python script do this work,
you need to run on the machine installed windows system and powerpoint
winxos 20230907
'''
import os
import win32com.client
from pypdf import PdfReader, PdfWriter

# 定义文件夹路径
folder_path = "F:\\Downloads\\中文版610本世界优秀绘本"
pps = win32com.client.Dispatch("PowerPoint.Application")
# 遍历文件夹中的所有ppt文件
for file in os.listdir(folder_path):
    if file.endswith(".pps") or file.endswith(".pptx") or file.endswith(".ppt"):
        # 获取文件的完整路径
        file_path = os.path.join(folder_path, file)
        # 获取文件的名称，不包括扩展名
        file_name = os.path.splitext(file)[0]
        # 创建一个pdf文件的路径，使用相同的名称
        pdf_path = os.path.join(folder_path, file_name + ".pdf")
        print(file_name)
        try:
            presentation = pps.Presentations.Open(file_path)
            presentation.SaveAs(pdf_path, 32) # 32代表pdf格式
            presentation.Close()
            
            # 打开pdf文件
            pdf = PdfReader(pdf_path)
            # 获取pdf文件的元数据对象
            # 创建一个新的元数据对象，将标题设置为文件名，作者设置为空
            new_metadata = PdfWriter()
            new_metadata.add_metadata({
                "/Title": file_name,
                "/Author": ""
            })
            # 将pdf文件的所有页面添加到新的元数据对象中
            for page in pdf.pages:
                new_metadata.add_page(page)
            # 用新的元数据对象覆盖原来的pdf文件
            with open(pdf_path, "wb") as f:
                new_metadata.write(f)
        except Exception as e:
            print(f"Error while processing {file}: {e}")
pps.Quit()
