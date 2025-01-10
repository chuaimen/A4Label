'''
有哪些名称需要统一

有图片 和 没图片
    信息录入情况

'''
import time
import tkinter as tk
import os
import PIL
from tkinter import PhotoImage
from PIL import Image, ImageTk
from tkinter import ttk
from productListEXCEL import A4LABELproductListRun
import datetime

class Label4():
    def __init__(self):
        #/Users/mac/Desktop/苹果系统/A4LabelFor4
        self.main_path = os.getcwd()
        temp_path = os.path.join(self.main_path, 'document')
        self.clientDocument_path = os.path.join(temp_path,'clientDocument')
        self.creatDocument_path = os.path.join(temp_path, 'creatnewExcelfile')
        self.creatDocument_path = os.path.join(self.creatDocument_path, 'A4LabelExcel.xlsx')
        self.select_client = ''
        self.pic_file_path = os.path.join(temp_path,'pinPicture')
        self.temp_pic_path = os.path.join(self.pic_file_path,'temppic')

        #替代图片 填补空的项的图片
        self.pic_path_for_blank = os.path.join(self.pic_file_path,'aa.png')


        self.pic_path_2 = os.path.join(self.pic_file_path, 'tt.jpeg')
        self.a = 0
        #在子-子窗口自动合成 选中文件路径
        self.temp_excel_cilen_path = '' #读取选中的客户文件

        self.child_child_select_photo = ''

        self.select_information = {
            'window1_pic': '', 'client1': '', 'productNumber1': '', 'productName1': '', 'time1': '',
            'window2_pic': '', 'client2': '', 'productNumber2': '', 'productName2': '', 'time2': '',
            'window3_pic': '', 'client3': '', 'productNumber3': '', 'productName3': '', 'time3': '',
            'window4_pic': '', 'client4': '', 'productNumber4': '', 'productName4': '', 'time4': '',
        }


#4 button
    def button_1(self):
        self.window = 'windows 1'
        self.open_child_window()
    def button_2(self):
        self.window = 'windows 2'
        self.open_child_window()
    def button_3(self):
        self.window = 'windows 3'
        self.open_child_window()
    def button_4(self):
        self.window = 'windows 4'
        self.open_child_window()

    def open_child_window(self):
        # 创建一个新的Toplevel窗口作为子窗口
        self.child_window = tk.Toplevel()
        self.child_window.title("子窗口")

        # 示例2：同时设置窗口大小和位置
        # 假设屏幕宽度为1920像素，高度为1080像素
        # 我们想要将窗口放置在屏幕中央
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 500
        window_height = 1080
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        #self.child_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.child_window.geometry(f"{window_width}x{window_height}")
        # 创建一个Scrollbar
        scrollbar = tk.Scrollbar(self.child_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 创建一个Listbox，并设置yscrollcommand为Scrollbar的set方法
        self.child_listbox = tk.Listbox(self.child_window, yscrollcommand=scrollbar.set)
        self.child_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

            # 在子窗口中添加一些控件（例如，一个标签）
        label = tk.Label(self.child_window, text=self.window)
        label.pack(pady=20)

        # 配置Scrollbar的command选项为Listbox的yview方法
        scrollbar.config(command=self.child_listbox.yview)

        #def add_item():
        for i in os.listdir(self.clientDocument_path):
            self.child_listbox.insert(tk.END, i)

        # 添加一个按钮来添加新项到Listbox中，以便可以看到滚动效果
        #add_button = tk.Button(self.child_window, text="添加项", command=add_item)
        #add_button.pack(side=tk.BOTTOM)



        add_button = tk.Button(self.child_window, text='打开', command=self.open_child_child_window)
        add_button.pack(side=tk.BOTTOM)

    #in 子窗口 set self.select_clien
    def get_selected_text(self):
        selected_indices = self.child_listbox.curselection()
        if selected_indices:
            # 假设我们只关心第一个选中的项
            index = selected_indices[0]
            # 根据索引获取选中的项
            selected_item = self.child_listbox.get(index)
            print(f"选中的项是: {selected_item}")
            self.select_client = selected_item
            self.temp_excel_cilen_path = os.path.join(self.clientDocument_path, self.select_client)


    def open_child_child_window(self):
        # 创建一个新的Toplevel窗口作为子窗口
        self.child_child_window = tk.Toplevel()
        self.child_child_window.title("子窗口")

        self.get_selected_text()
        #保存选中文件的图片


        self.A4LABEL_Excel = A4LABELproductListRun.ExcelEdit()
        # 保存选中文件的图片
        self.A4LABEL_Excel.get_image_for_save(
            self.temp_excel_cilen_path,self.temp_pic_path)

        maxrow, maxcolumn = self.A4LABEL_Excel.A4LABEL_max_row_column(self.temp_excel_cilen_path)

        # 示例2：同时设置窗口大小和位置
        # 假设屏幕宽度为1920像素，高度为1080像素
        # 我们想要将窗口放置在屏幕中央
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 800
        window_height = 1080
        x = (screen_width // 2) - (window_width // 2) - 50
        y = (screen_height // 2) - (window_height // 2)
        self.child_child_window.geometry(f"{window_width}x{window_height}")

        self.child_child_frame_1 = tk.Frame(self.child_child_window)
        self.child_child_frame_2 = tk.Frame(self.child_child_window)

        self.child_child_frame_1.pack(side=tk.LEFT, expand=True, fill=tk.Y)
        self.child_child_frame_2.pack(side=tk.RIGHT, expand=True, fill=tk.Y)

        # 创建一个树状视图
        tree = ttk.Treeview(self.child_child_frame_1, columns=( '物料编号', '名称'))
        tree.heading("物料编号", text="物料编号")
        tree.heading("名称", text="名称")


        tree.column('物料编号', width=100, minwidth=150, stretch=tk.NO)
        tree.column('名称', width=100, minwidth=270, stretch=tk.NO)

        style = ttk.Style()
        style.configure('Treeview', rowheight=70)

        # 读取第一行
        temp_client_title_data = self.A4LABEL_Excel.A4LABEL_get_title_information(self.temp_excel_cilen_path)
        #print(temp_client_title_data)
        try:
            for i in range(len(temp_client_title_data)):
                if temp_client_title_data[i] == '图片':
                    self.pic_column = i + 1
                elif temp_client_title_data[i] == '物料编号':
                    self.product_number = i + 1
                elif temp_client_title_data[i] == '产品名称':
                    self.product_name = i + 1
        except AttributeError as e:
            print('in Label open_child_child_window line 157', e)


        #图片
        '''image = Image.open(self.pic_path)  # 替换为你的图片路径
        image = image.resize((100, 50), Image.LANCZOS)
        self.photo = ImageTk.PhotoImage(image)
        tree.insert('', i, image=self.temp_pic[i], values=['2', '3'])'''

        self.temp_pic = {}
        for i in os.listdir(self.temp_pic_path):
            pic_path = os.path.join(self.temp_pic_path, i)
            image = Image.open(pic_path)  # 替换为你的图片路径
            image = image.resize((100, 50), Image.LANCZOS)

            self.temp_pic[i[:-4]] = ImageTk.PhotoImage(image)


        for i in range(2, maxrow+1):
            if self.temp_pic.get(str(i),'not found') == 'not found':

                image = Image.open(self.pic_path_for_blank)  # 替换为你的图片路径
                image = image.resize((100, 50), Image.LANCZOS)
                self.temp_pic[str(i)] = ImageTk.PhotoImage(image)
                image.save(os.path.join(self.temp_pic_path,str(i)+'.png'))


            try:
                tree.insert('', i, image=self.temp_pic[str(i)],text=str(i), values=[
                    self.A4LABEL_Excel.A4LABEL_get_row_information(self.temp_excel_cilen_path,i,self.product_number),
                    self.A4LABEL_Excel.A4LABEL_get_row_information(self.temp_excel_cilen_path,i,self.product_name)
                ])

            except IOError as e:
                print(e)

        def on_select(event):
            selected_item = tree.selection()[0]  # 获取选中项的IID（项标识符）
            item = tree.item(selected_item)  # 使用IID获取项的信息

            self.select_pic_row = item['text']
            self.select_product_number = self.A4LABEL_Excel.A4LABEL_get_row_information(self.temp_excel_cilen_path,int(self.select_pic_row),self.product_number)
            self.select_product_name = self.A4LABEL_Excel.A4LABEL_get_row_information(self.temp_excel_cilen_path,int(self.select_pic_row),self.product_name)

            print(f"选中项的文本: {item.items()}") #选中项的文本: ['pyimage10']
            #self.child_child_select_photo = self.temp_pic[str(i)]

            self.child_child_select_photo = os.path.join(self.temp_pic_path, self.select_pic_row+'.png')
            if os.path.exists(self.child_child_select_photo):
                print('self.child_child_select_photo >>> ',self.child_child_select_photo)
            else:
                self.child_child_select_photo = ''
                print('self.child_child_select_photo >>> not exists ')



            print(self.select_pic_row, self.select_product_number,self.select_product_name)


        tree.bind("<<TreeviewSelect>>", on_select)
        tree.pack(expand=True, fill='both')

        label = tk.Label(self.child_child_frame_2, text=self.window)
        add_button = tk.Button(self.child_child_frame_2, text="选中close", command=self.button_child_child_command)

        label.pack(side=tk.TOP)
        add_button.pack(side=tk.BOTTOM)



    #关闭子_子窗口 实现的功能
    def button_child_child_command(self):
        self.child_child_window.destroy()
        self.child_window.destroy()

        '''self.select_information = {
            'window1_pic':'','client1':'','productNumber1':'','productName1':'','time1':'',
            'window2_pic':'', 'client2': '', 'productNumber2': '', 'productName2': '', 'time2': '',
            'window3_pic': '', 'client3': '', 'productNumber3': '', 'productName3': '', 'time3': '',
            'window4_pic': '', 'client4': '', 'productNumber4': '', 'productName4': '', 'time4': '',
        }'''

        try:

            if self.window == 'windows 1':
                if self.child_child_select_photo != '':
                    image_1 = Image.open(self.child_child_select_photo)
                    image_1 = image_1.resize((320, 170), Image.LANCZOS)
                    self.photo_1 = ImageTk.PhotoImage(image_1)

                    image_1.save(os.path.join(self.pic_file_path, 'window1.png'))
                    self.select_information['window1_pic'] = os.path.join(self.pic_file_path, 'window1.png')

                    self.button1.config(image=self.photo_1, height=170, )
                    self.button1.grid(row=0, column=0, sticky="nsew", columnspan=2)

                    self.entry11.delete(0, tk.END)  # 删除Entry中的所有内容
                    self.entry11.insert(0, self.select_client[:-5])  # 在Entry的开始位置插入新文本
                    self.select_information['client1']= self.select_client[:-5]

                    self.entry12.delete(0, tk.END)  # 删除Entry中的所有内容
                    self.entry12.insert(0, self.select_product_number)  # 在Entry的开始位置插入新文本
                    self.select_information['productNumber1'] = self.select_product_number[0]

                    self.entry13.delete(0, tk.END)
                    self.entry13.insert(0, self.select_product_name)
                    self.select_information['productName1'] = self.select_product_name[0]

                    self.entry14.delete(0, tk.END)
                    self.entry14.insert(0, str(datetime.date.today()))
                    self.select_information['time1']= str(datetime.date.today())

                    print(self.select_information)


            elif self.window == 'windows 2':
                if self.child_child_select_photo != '':
                    image_2 = Image.open(self.child_child_select_photo)  # 替换为你的图片路径
                    image_2 = image_2.resize((320, 170), Image.LANCZOS)
                    self.photo_2 = ImageTk.PhotoImage(image_2)

                    image_2.save(os.path.join(self.pic_file_path, 'window2.png'))
                    self.select_information['window2_pic'] = os.path.join(self.pic_file_path, 'window2.png')

                    self.button2.config(image=self.photo_2, height=170)
                    self.button2.grid(row=0, column=0, sticky="nsew", columnspan=2)

                    self.entry21.delete(0, tk.END)  # 删除Entry中的所有内容
                    self.entry21.insert(0, self.select_client[:-5])  # 在Entry的开始位置插入新文本
                    self.select_information['client2'] = self.select_client[:-5]

                    self.entry22.delete(0, tk.END)  # 删除Entry中的所有内容
                    self.entry22.insert(0, self.select_product_number)  # 在Entry的开始位置插入新文本
                    self.select_information['productNumber2'] = self.select_product_number[0]

                    self.entry23.delete(0, tk.END)
                    self.entry23.insert(0, self.select_product_name)
                    self.select_information['productName2'] = self.select_product_name[0]

                    self.entry24.delete(0, tk.END)
                    self.entry24.insert(0, str(datetime.date.today()))
                    self.select_information['time2'] = str(datetime.date.today())

                    print(self.select_information)

            elif self.window == 'windows 3':
                if self.child_child_select_photo != '':
                    image_3 = Image.open(self.child_child_select_photo)  # 替换为你的图片路径
                    image_3 = image_3.resize((320, 170), Image.LANCZOS)
                    self.photo_3 = ImageTk.PhotoImage(image_3)

                    image_3.save(os.path.join(self.pic_file_path, 'window3.png'))
                    self.select_information['window3_pic'] = os.path.join(self.pic_file_path, 'window3.png')

                    self.button3.config(image=self.photo_3, height=170)
                    self.button3.grid(row=0, column=0, sticky="nsew", columnspan=2)


                    self.entry31.delete(0, tk.END)  # 删除Entry中的所有内容
                    self.entry31.insert(0, self.select_client[:-5])  # 在Entry的开始位置插入新文本
                    self.select_information['client3'] = self.select_client[:-5]

                    self.entry32.delete(0, tk.END)  # 删除Entry中的所有内容
                    self.entry32.insert(0, self.select_product_number)  # 在Entry的开始位置插入新文本
                    self.select_information['productNumber3'] = self.select_product_number[0]

                    self.entry33.delete(0, tk.END)
                    self.entry33.insert(0, self.select_product_name)
                    self.select_information['productName3'] = self.select_product_name[0]

                    self.entry34.delete(0, tk.END)
                    self.entry34.insert(0, str(datetime.date.today()))
                    self.select_information['time3'] = str(datetime.date.today())

                    print(self.select_information)


            elif self.window == 'windows 4':
                if self.child_child_select_photo != '':
                    image_4 = Image.open(self.child_child_select_photo)  # 替换为你的图片路径
                    image_4 = image_4.resize((320, 170), Image.LANCZOS)
                    self.photo_4 = ImageTk.PhotoImage(image_4)

                    image_4.save(os.path.join(self.pic_file_path, 'window4.png'))
                    self.select_information['window4_pic'] = os.path.join(self.pic_file_path, 'window4.png')

                    self.button4.config(image=self.photo_4, height=170)
                    self.button4.grid(row=0, column=0, sticky="nsew", columnspan=2)

                    self.entry41.delete(0, tk.END)  # 删除Entry中的所有内容
                    self.entry41.insert(0, self.select_client[:-5])  # 在Entry的开始位置插入新文本
                    self.select_information['client4'] = self.select_client[:-5]

                    self.entry42.delete(0, tk.END)  # 删除Entry中的所有内容
                    self.entry42.insert(0, self.select_product_number)  # 在Entry的开始位置插入新文本
                    self.select_information['productNumber4'] = self.select_product_number[0]

                    self.entry43.delete(0, tk.END)
                    self.entry43.insert(0, self.select_product_name)
                    self.select_information['productName4'] = self.select_product_name[0]

                    self.entry44.delete(0, tk.END)
                    self.entry44.insert(0, str(datetime.date.today()))
                    self.select_information['time4'] = str(datetime.date.today())

                    print(self.select_information)


        except Exception as e:
            print(e)
    def information_puase_4(self):
        image = Image.open(self.child_child_select_photo)  # 替换为你的图片路径
        image = image.resize((320, 170), Image.LANCZOS)
        self.photo_4 = ImageTk.PhotoImage(image)

        image.save(os.path.join(self.pic_file_path, 'window2.png'))
        image.save(os.path.join(self.pic_file_path, 'window3.png'))
        image.save(os.path.join(self.pic_file_path, 'window4.png'))

        self.select_information['window2_pic'] = os.path.join(self.pic_file_path, 'window2.png')
        self.select_information['window3_pic'] = os.path.join(self.pic_file_path, 'window3.png')
        self.select_information['window4_pic'] = os.path.join(self.pic_file_path, 'window4.png')

        image_2 = Image.open(self.select_information['window2_pic'])  # 替换为你的图片路径
        image_2 = image_2.resize((320, 170), Image.LANCZOS)
        self.photo_2 = ImageTk.PhotoImage(image_2)
        image_3 = Image.open(self.select_information['window3_pic'])  # 替换为你的图片路径
        image_3 = image_3.resize((320, 170), Image.LANCZOS)
        self.photo_3 = ImageTk.PhotoImage(image_3)
        image_4 = Image.open(self.select_information['window4_pic'])  # 替换为你的图片路径
        image_4 = image_4.resize((320, 170), Image.LANCZOS)
        self.photo_4 = ImageTk.PhotoImage(image_4)

        self.button2.config(image=self.photo_2, height=170)
        self.button2.grid(row=0, column=0, sticky="nsew", columnspan=2)
        self.button3.config(image=self.photo_3, height=170)
        self.button3.grid(row=0, column=0, sticky="nsew", columnspan=2)
        self.button4.config(image=self.photo_4, height=170)
        self.button4.grid(row=0, column=0, sticky="nsew", columnspan=2)

        self.select_information['client2'] = self.select_information['client1']
        self.entry21.delete(0, tk.END)  # 删除Entry中的所有内容
        self.entry21.insert(0, self.select_information['client2'])  # 在Entry的开始位置插入新文本
        self.select_information['client3'] = self.select_information['client1']
        self.entry31.delete(0, tk.END)  # 删除Entry中的所有内容
        self.entry31.insert(0, self.select_information['client3'])  # 在Entry的开始位置插入新文本
        self.select_information['client4'] = self.select_information['client1']
        self.entry41.delete(0, tk.END)  # 删除Entry中的所有内容
        self.entry41.insert(0, self.select_information['client4'])  # 在Entry的开始位置插入新文本

        self.select_information['productNumber2'] = self.select_information['productNumber1']
        self.entry22.delete(0, tk.END)  # 删除Entry中的所有内容
        self.entry22.insert(0, self.select_information['productNumber2'])  # 在Entry的开始位置插入新文本
        self.select_information['productNumber3'] = self.select_information['productNumber1']
        self.entry32.delete(0, tk.END)  # 删除Entry中的所有内容
        self.entry32.insert(0, self.select_information['productNumber3'])  # 在Entry的开始位置插入新文本
        self.select_information['productNumber4'] = self.select_information['productNumber1']
        self.entry42.delete(0, tk.END)  # 删除Entry中的所有内容
        self.entry42.insert(0, self.select_information['productNumber4'])  # 在Entry的开始位置插入新文本

        self.select_information['productName2'] = self.select_information['productName1']
        self.entry23.delete(0, tk.END)
        self.entry23.insert(0, self.select_information['productName2'] )
        self.select_information['productName3'] = self.select_information['productName1']
        self.entry33.delete(0, tk.END)
        self.entry33.insert(0, self.select_information['productName3'] )
        self.select_information['productName4'] = self.select_information['productName1']
        self.entry43.delete(0, tk.END)
        self.entry43.insert(0, self.select_information['productName4'] )


        self.entry24.delete(0, tk.END)
        self.entry24.insert(0, str(datetime.date.today()))
        self.select_information['time2'] = str(datetime.date.today())
        self.entry34.delete(0, tk.END)
        self.entry34.insert(0, str(datetime.date.today()))
        self.select_information['time3'] = str(datetime.date.today())
        self.entry44.delete(0, tk.END)
        self.entry44.insert(0, str(datetime.date.today()))
        self.select_information['time4'] = str(datetime.date.today())

        print(self.select_information)
    def buttton_open_insert_to_excel(self):
        self.A4LABEL_Excel.A4_LABEL_init_Excel_file(self.creatDocument_path)

        self.A4LABEL_Excel.A4_LABEL_insert_Excel_file(self.creatDocument_path,self.select_information)

        cmd = 'start ' + self.creatDocument_path
        os.system(cmd)
    def LabelRun(self):
            # 创建一个主窗口
            self.root = tk.Tk()
            self.root.title("Grid布局嵌套示例")

            # 示例2：同时设置窗口大小和位置
            # 假设屏幕宽度为1920像素，高度为1080像素
            # 我们想要将窗口放置在屏幕中央
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            window_width = 1080
            window_height = 780
            x = (screen_width // 2) - (window_width // 2)
            y = (screen_height // 2) - (window_height // 2)
            self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

            # 创建一个Frame作为内部容器，并设置其布局管理器为Grid
            self.inner_frame = tk.Frame(self.root)
            inner_frame_2 = tk.Frame(self.root)
            inner_frame_3 = tk.Frame(self.root)
            inner_frame_4 = tk.Frame(self.root)
            inner_frame_5 = tk.Frame(self.root)
            inner_frame_6 = tk.Frame(self.root)
            self.inner_frame.grid(row=0, column=0, sticky="nsew")  # 占据整个主窗口，并拉伸以填充空间
            inner_frame_2.grid(row=0, column=1, sticky="nsew")
            inner_frame_3.grid(row=1, column=0, sticky="nsew")
            inner_frame_4.grid(row=1, column=1, sticky="nsew")
            inner_frame_5.grid(row=2, column=0, sticky="nsew")
            inner_frame_6.grid(row=2, column=1, sticky="nsew")


            # 在内部Frame中创建控件
            # 第1.1个标签
            # 创建一个按钮，并设置其初始宽度和高度
            self.button1 = tk.Button(self.inner_frame, text="图片", width=10, height=10, command=self.button_1)
            self.button1.grid(row=0, column=0, sticky="nsew",columnspan=2)  # 使用sticky使按钮填充整个单元格
            # 第1.2个标签
            label1 = tk.Label(self.inner_frame, text="厂家", width=10, height=1)
            label1.grid(row=1, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry11 = tk.Entry(self.inner_frame, width=30)  # width参数设置输入框的宽度
            self.entry11.grid(row=1, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第1.3个标签，跨越两列
            label2 = tk.Label(self.inner_frame, text="产品编码", width=10, height=1)
            label2.grid(row=2, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry12 = tk.Entry(self.inner_frame, width=30)  # width参数设置输入框的宽度
            self.entry12.grid(row=2, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第1.3个标签，跨越两列
            label3 = tk.Label(self.inner_frame, text="产品名称", width=10, height=1)
            label3.grid(row=3, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry13 = tk.Entry(self.inner_frame, width=30)  # width参数设置输入框的宽度
            self.entry13.grid(row=3, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第1.3个标签，跨越两列
            label4 = tk.Label(self.inner_frame, text="生产日期", width=10, height=1)
            label4.grid(row=4, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry14 = tk.Entry(self.inner_frame, width=30)  # width参数设置输入框的宽度
            self.entry14.grid(row=4, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距

            # 第2.1个标签
            # 创建一个按钮，并设置其初始宽度和高度
            self.button2 = tk.Button(inner_frame_2, text="2图片", width=10, height=10, command=self.button_2)
            self.button2.grid(row=0, column=0, sticky="nsew",columnspan=2)  # 使用sticky使按钮填充整个单元格
            # 第2.2个标签
            label1 = tk.Label(inner_frame_2, text="厂家", width=10, height=1)
            label1.grid(row=1, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry21 = tk.Entry(inner_frame_2, width=30)  # width参数设置输入框的宽度
            self.entry21.grid(row=1, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第2.3个标签，跨越两列
            label2 = tk.Label(inner_frame_2, text="产品编码", width=10, height=1)
            label2.grid(row=2, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry22 = tk.Entry(inner_frame_2, width=30)  # width参数设置输入框的宽度
            self.entry22.grid(row=2, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第2.3个标签，跨越两列
            label3 = tk.Label(inner_frame_2, text="产品名称", width=10, height=1)
            label3.grid(row=3, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry23 = tk.Entry(inner_frame_2, width=30)  # width参数设置输入框的宽度
            self.entry23.grid(row=3, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第2.3个标签，跨越两列
            label4 = tk.Label(inner_frame_2, text="生产日期", width=10, height=1)
            label4.grid(row=4, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry24 = tk.Entry(inner_frame_2, width=30)  # width参数设置输入框的宽度
            self.entry24.grid(row=4, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距

        

            # 第3.1个标签
            # 创建一个按钮，并设置其初始宽度和高度
            self.button3 = tk.Button(inner_frame_3, text="3图片", width=10, height=10, command=self.button_3)
            self.button3.grid(row=0, column=0, sticky="nsew",columnspan=2)  # 使用sticky使按钮填充整个单元格
            # 第3.2个标签
            label1 = tk.Label(inner_frame_3, text="厂家", width=10, height=1)
            label1.grid(row=1, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry31 = tk.Entry(inner_frame_3, width=30)  # width参数设置输入框的宽度
            self.entry31.grid(row=1, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第3.3个标签，跨越两列
            label2 = tk.Label(inner_frame_3, text="产品编码", width=10, height=1)
            label2.grid(row=2, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry32 = tk.Entry(inner_frame_3, width=30)  # width参数设置输入框的宽度
            self.entry32.grid(row=2, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第3.3个标签，跨越两列
            label3 = tk.Label(inner_frame_3, text="产品名称", width=10, height=1)
            label3.grid(row=3, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry33 = tk.Entry(inner_frame_3, width=30)  # width参数设置输入框的宽度
            self.entry33.grid(row=3, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第3.3个标签，跨越两列
            label4 = tk.Label(inner_frame_3, text="生产日期", width=10, height=1)
            label4.grid(row=4, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry34 = tk.Entry(inner_frame_3, width=30)  # width参数设置输入框的宽度
            self.entry34.grid(row=4, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距


            # 第4.1个标签
            # 创建一个按钮，并设置其初始宽度和高度
            self.button4 = tk.Button(inner_frame_4, text="4图片", width=10, height=10, command=self.button_4)
            self.button4.grid(row=0, column=0, sticky="nsew",columnspan=2)  # 使用sticky使按钮填充整个单元格
            # 第4.2个标签
            label1 = tk.Label(inner_frame_4, text="厂家", width=10, height=1)
            label1.grid(row=1, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry41 = tk.Entry(inner_frame_4, width=30)  # width参数设置输入框的宽度
            self.entry41.grid(row=1, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第4.3个标签，跨越两列
            label2 = tk.Label(inner_frame_4, text="产品编码", width=10, height=1)
            label2.grid(row=2, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry42 = tk.Entry(inner_frame_4, width=30)  # width参数设置输入框的宽度
            self.entry42.grid(row=2, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第4.3个标签，跨越两列
            label3 = tk.Label(inner_frame_4, text="产品名称", width=10, height=1)
            label3.grid(row=3, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry43 = tk.Entry(inner_frame_4, width=30)  # width参数设置输入框的宽度
            self.entry43.grid(row=3, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距
            # 第4.3个标签，跨越两列
            label4 = tk.Label(inner_frame_4, text="生产日期", width=10, height=1)
            label4.grid(row=4, column=0, padx=10, pady=10)
            # 创建一个Entry输入框
            self.entry44 = tk.Entry(inner_frame_4, width=30)  # width参数设置输入框的宽度
            self.entry44.grid(row=4, column=1, padx=10, pady=10)  # 将Entry添加到窗口中，并设置上下外边距

            # 第5.1个标签
            # 创建一个按钮，并设置其初始宽度和高度
            self.button5 = tk.Button(inner_frame_5, text="打 开", command=self.buttton_open_insert_to_excel)
            self.button5.grid(row=0, column=0,sticky="nsew" ,columnspan=2)  # 使用sticky使按钮填充整个单元格

            # 第5.2个标签
            # 创建一个按钮，*4 x4 x4
            self.button6 = tk.Button(inner_frame_6, text="X4X4", command=self.information_puase_4)
            self.button6.grid(row=0,column=1, sticky="nsew",columnspan=2)  # 使用sticky使按钮填充整个单元格

            # 配置内部Frame的列权重，以便其子控件可以拉伸以填充空间
            #inner_frame.grid_columnconfigure(0, weight=1)
            #inner_frame_2.grid_columnconfigure(0, weight=1)
            #inner_frame_3.grid_columnconfigure(0, weight=1)
            inner_frame_5.grid_columnconfigure(1, weight=1)
            inner_frame_6.grid_columnconfigure(1, weight=1)

            # 配置主窗口的列权重（在这个例子中其实不需要，因为只有一个内部Frame）
            # 但如果你打算在主窗口中直接放置更多控件，并希望它们拉伸以填充空间，那么就需要这样做
            self.root.grid_columnconfigure(0, weight=1)
            self.root.grid_columnconfigure(1, weight=1)
            # 启动事件循环
            self.root.mainloop()

if __name__ == "__main__":

    temp = Label4()
    temp.LabelRun()

