'''
ppt,pps,pptx files to pdf
many ebooks is ppt file, calibre can't manage,
so this python script do this work,
you need to run on the machine installed windows system and powerpoint
winxos 20230907
'''
import os
from pypdf import PdfReader, PdfWriter
import fitz  # 导入 PyMuPDF 库
# 定义文件夹路径
folder_path = "G:\\EBOOKS"
idx = 0
fs  = []
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith(".pdf"): # 只处理 pdf 文件
            file_path = os.path.join(root, file) # 获取完整的文件路径
            fs.append((file_path,file))
skip = 5770
for file_path,file in fs[skip:]:
    idx +=1
    try:      
        file_name = os.path.splitext(file)[0]
        info = file_name.split(' - ')
        doc = fitz.open(file_path)
        metadata = doc.metadata
        if metadata == None or metadata['title'] != info[0] or metadata['author'] != info[1]:
            print(idx+skip,file_name)
            metadata["title"] = info[0]
            metadata["author"] = info[1]
            doc.set_metadata(metadata)
            doc.save(file_path,incremental=True, encryption=0)
    except Exception as e:
        print(f"Error while processing {file}: {e}")
