import tkinter as tk
from tkinter import ttk
from tkinter import PhotoImage
import tkinter as tk
import os
import PIL
from tkinter import PhotoImage
from PIL import Image, ImageTk
from tkinter import ttk
from productListEXCEL import A4LABELproductListRun


#########

'''def insert_image(tree, parent, image):
    # 在树状视图中插入图片
    iid = tree.insert(parent, 'end', image=image)
    return iid'''


def main():
    child_child_window = tk.Tk()
    child_child_window.title("Tree View with Images")

    # 示例2：同时设置窗口大小和位置
    # 假设屏幕宽度为1920像素，高度为1080像素
    # 我们想要将窗口放置在屏幕中央
    screen_width = child_child_window.winfo_screenwidth()
    screen_height = child_child_window.winfo_screenheight()
    window_width = 800
    window_height = 1080
    x = (screen_width // 2) - (window_width // 2) - 50
    y = (screen_height // 2) - (window_height // 2)
    child_child_window.geometry(f"{window_width}x{window_height}")

    child_child_frame_1 = tk.Frame(child_child_window)
    child_child_frame_2 = tk.Frame(child_child_window)

    child_child_frame_1.pack(side=tk.LEFT, expand=True, fill=tk.Y)
    child_child_frame_2.pack(side=tk.RIGHT, expand=True, fill=tk.Y)

    # 创建一个树状视图
    tree = ttk.Treeview(child_child_frame_1, columns=('图片','物料编号','名称'))
    tree.heading("图片", text="图片")
    tree.heading("物料编号", text="物料编号")
    tree.heading("名称", text="名称")

    tree.column('图片', width=100, minwidth=270, stretch=tk.NO)
    tree.column('物料编号', width=100, minwidth=150, stretch=tk.NO)
    tree.column('名称', width=100, minwidth=270, stretch=tk.NO)

    # 创建图片
    # 加载图片
    image_path = 'C:\\Users\\hello\\Desktop\\A4LabelFor4\\document\\pinPicture\\aa.jpg'
    image = Image.open(image_path)
    image = image.resize((100, 50), Image.LANCZOS)
    image = ImageTk.PhotoImage(image)
    # 插入图片到树状视图的根节点

    tree.insert('',tk.END,image=image,values=['2','3'])
    tree.insert('', tk.END, image=image, values=['2', '3'])
    tree.insert('', tk.END, image=image, values=['2', '3'])
    tree.insert('', tk.END, image=image, values=['2', '3'])
    tree.pack(expand=True, fill='both')

    label = tk.Label(child_child_frame_2, text='111')
    add_button = tk.Button(child_child_frame_2, text="选中close")

    # tree.pack()
    label.pack(side=tk.TOP)
    add_button.pack(side=tk.BOTTOM)

    child_child_window.mainloop()


if __name__ == '__main__':
    main()





