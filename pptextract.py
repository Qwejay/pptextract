import os
import shutil
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import zipfile
import mimetypes
import win32com.client as win32

def convert_ppt_to_pptx(ppt_path):
    """
    将PPT文件转换为PPTX文件
    :param ppt_path: PPT文件的路径
    :return: 转换后的PPTX文件路径，如果转换失败则返回None
    """
    try:
        powerpoint = win32.gencache.EnsureDispatch('Powerpoint.Application')
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        pptx_path = os.path.splitext(ppt_path)[0] + '.pptx'
        presentation.SaveAs(pptx_path, 24)  # 24 represents the PPTX file format
        presentation.Close()
        powerpoint.Quit()
        return pptx_path
    except Exception as e:
        messagebox.showerror("错误", f"无法转换PPT文件: {e}")
        return None

def extract_media(ppt_path):
    """
    从PPTX文件中提取媒体文件（图片和音频）
    :param ppt_path: PPT或PPTX文件的路径
    :return: 提取的媒体文件列表
    """
    try:
        if ppt_path.endswith('.ppt'):
            ppt_path = convert_ppt_to_pptx(ppt_path)
            if not ppt_path:
                return []
        
        ppt_name = os.path.splitext(os.path.basename(ppt_path))[0]
        current_dir = os.path.dirname(ppt_path)
        media_folder = os.path.join(current_dir, ppt_name)
        os.makedirs(media_folder, exist_ok=True)
        
        image_folder = os.path.join(media_folder, '图片')
        audio_folder = os.path.join(media_folder, '音乐')
        os.makedirs(image_folder, exist_ok=True)
        os.makedirs(audio_folder, exist_ok=True)
        
        with zipfile.ZipFile(ppt_path, 'r') as zip_ref:
            media_files = [f for f in zip_ref.namelist() if f.startswith('ppt/media/')]
            for file in media_files:
                file_type, _ = mimetypes.guess_type(file)
                if file_type and file_type.startswith('image'):
                    zip_ref.extract(file, image_folder)
                elif file_type and file_type.startswith('audio'):
                    zip_ref.extract(file, audio_folder)
        
        # Move files from nested 'ppt/media' directories to the root of '图片' and '音乐' folders
        for root, dirs, files in os.walk(image_folder):
            for file in files:
                shutil.move(os.path.join(root, file), os.path.join(image_folder, file))
        for root, dirs, files in os.walk(audio_folder):
            for file in files:
                shutil.move(os.path.join(root, file), os.path.join(audio_folder, file))
        
        # Remove empty 'ppt/media' directories
        shutil.rmtree(os.path.join(image_folder, 'ppt'), ignore_errors=True)
        shutil.rmtree(os.path.join(audio_folder, 'ppt'), ignore_errors=True)
        
        # Rename files to avoid None in filenames
        for root, dirs, files in os.walk(image_folder):
            for file in files:
                if 'None' in file:
                    new_file = file.replace('None', '')
                    os.rename(os.path.join(root, file), os.path.join(root, new_file))
        for root, dirs, files in os.walk(audio_folder):
            for file in files:
                if 'None' in file:
                    new_file = file.replace('None', '')
                    os.rename(os.path.join(root, file), os.path.join(root, new_file))
        
        return os.listdir(image_folder) + os.listdir(audio_folder)
    except Exception as e:
        messagebox.showerror("错误", f"无法提取媒体文件: {e}")
        return []

def on_drop(event):
    """
    处理文件拖放事件
    :param event: 拖放事件
    """
    file_path = event.data
    if file_path:
        # 处理文件路径，去除可能的特殊字符
        file_path = file_path.strip('{}')
        file_path = file_path.replace('\\', '/')
        if file_path.startswith('file://'):
            file_path = file_path[7:]
        file_path = os.path.normpath(file_path)
        
        if os.path.isfile(file_path):
            media_files = extract_media(file_path)
            if media_files:
                messagebox.showinfo("成功", f"已提取 {len(media_files)} 个媒体文件")
            else:
                messagebox.showinfo("信息", "未在PPTX文件中找到媒体文件")
        else:
            messagebox.showerror("错误", "无效的文件路径")

def main():
    """
    主函数，创建并运行Tkinter窗口
    """
    root = TkinterDnD.Tk()
    root.title("PPT媒体提取器 v1.1")
    root.geometry("300x100")  # 设置窗口大小

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop)

    label = tk.Label(root, text="将PPT文件拖放到此处开始提取媒体文件\nBy Qwejay")
    label.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
