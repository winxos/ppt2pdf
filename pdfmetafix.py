'''
ppt,pps,pptx files to pdf
many ebooks is ppt file, calibre can't manage,
so this python script do this work,
you need to run on the machine installed windows system and powerpoint
winxos 20230907
'''
import os
from pypdf import PdfReader, PdfWriter

# 定义文件夹路径
folder_path = "G:\\EBOOKS"
idx = 0
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith(".pdf"): # 只处理 pdf 文件
            idx +=1
            if idx < 3184:
                continue
            file_path = os.path.join(root, file) # 获取完整的文件路径
            try:      
                file_name = os.path.splitext(file)[0]
                info = file_name.split(' - ')
                # 打开pdf文件
                pdf = PdfReader(file_path)
                if pdf.metadata == None or pdf.metadata.title != info[0] or pdf.metadata.author != info[1]:
                    print(idx, file_path)
                    new_metadata = PdfWriter()
                    new_metadata.add_metadata({
                        "/Title": info[0],
                        "/Author": info[1]
                    })
                    # 将pdf文件的所有页面添加到新的元数据对象中
                    new_metadata.append_pages_from_reader(pdf)
                    
                    # 用新的元数据对象覆盖原来的pdf文件
                    with open(file_path, "wb") as f:
                        new_metadata.write(f)
            except Exception as e:
                print(f"Error while processing {file}: {e}")
