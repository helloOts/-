import tkinter as tk
from tkinter import filedialog
from PIL import Image
import os
import math


def resize_image(img, target_width, target_height):
    width, height = img.size
    ratio = min(target_width/width, target_height/height)
    new_width, new_height = int(width*ratio), int(height*ratio)
    return img.resize((new_width, new_height), Image.ANTIALIAS)

def combine_images(input_folder, output_folder):
    images = [Image.open(input_folder + '/' + img) for img in os.listdir(input_folder) if img.endswith(('.png', '.jpg', '.jpeg'))]
    a4_dims = (2480, 3508)  # A4 dimensions in pixels at 300dpi

    num_pages = math.ceil(len(images) / 6)
    for page in range(num_pages):
        combined = Image.new('RGB', a4_dims, (255, 255, 255))  # Create a white background
        for i, img in enumerate(images[page*6:(page+1)*6]):
            img = resize_image(img, a4_dims[0]/2, a4_dims[1]/3)
            x_offset = (i % 2) * int(a4_dims[0]/2)
            y_offset = (i // 2) * int(a4_dims[1]/3)
            combined.paste(img, (x_offset, y_offset))
        combined.save(f'{output_folder}/output_page_{page+1}.jpg')
    return num_pages  # 返回页数

def select_input_folder():
    input_folder = filedialog.askdirectory(title="Select Input Folder")
    input_entry.delete(0, tk.END)
    input_entry.insert(0, input_folder)

def select_output_folder():
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_folder)

def generate():
    input_folder = input_entry.get()
    output_folder = output_entry.get()
    if input_folder and output_folder:
        num_pages = combine_images(input_folder, output_folder)  # 接收返回的页数
        tk.Label(root, text=f"Images Combined Successfully! {num_pages} page(s) created.").pack()
# 创建主窗口
root = tk.Tk()
root.title("Image Combiner")

# 创建输入文件夹部分
tk.Label(root, text="Input Folder:").pack()
input_entry = tk.Entry(root, width=50)
input_entry.pack()
tk.Button(root, text="Browse", command=select_input_folder).pack()

# 创建输出文件夹部分
tk.Label(root, text="Output Folder:").pack()
output_entry = tk.Entry(root, width=50)
output_entry.pack()
tk.Button(root, text="Browse", command=select_output_folder).pack()

# 创建生成按钮
tk.Button(root, text="Generate", command=generate).pack()

# 运行主循环
root.mainloop()
