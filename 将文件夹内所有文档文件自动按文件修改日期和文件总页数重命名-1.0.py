import os
import datetime
import win32com.client
from PyPDF2 import PdfReader
import pythoncom
from pptx import Presentation
import tkinter as tk
from tkinter import filedialog, messagebox


def get_file_modified_date(file_path):
    mtime = os.path.getmtime(file_path)
    return datetime.datetime.fromtimestamp(mtime).strftime('%Y%m%d')


def get_pdf_page_count(file_path):
    try:
        with open(file_path, 'rb') as f:
            return len(PdfReader(f).pages)
    except Exception:
        return 0


def create_office_app(app_names):
    """智能创建办公应用实例，支持多名称尝试"""
    for name in app_names:
        try:
            app = win32com.client.Dispatch(name)
            app.Visible = False
            return app
        except Exception:
            continue
    raise Exception(f"未找到可用的办公组件，尝试过的名称: {app_names}")


def create_ppt_app():
    """
    尝试创建 PowerPoint 或 WPS Presentation 应用，
    支持 "PowerPoint.Application", "KWPP.Application", "Wpp.Application"
    同时关闭弹窗
    """
    app_names = ["PowerPoint.Application", "KWPP.Application", "Wpp.Application"]
    for name in app_names:
        try:
            app = win32com.client.Dispatch(name)
            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            return app
        except Exception:
            continue
    raise Exception("未找到可用的演示文稿应用程序")


def get_doc_page_count(file_path):
    try:
        pythoncom.CoInitialize()
        # 优先级：Word > WPS 文字
        word_apps = ['Word.Application', 'Kw.Application', 'Wps.Application']
        app = create_office_app(word_apps)
        
        doc = app.Documents.Open(os.path.abspath(file_path))
        doc.Repaginate()
        
        try:
            count = doc.ComputeStatistics(2)  # 标准方法
        except Exception:
            count = doc.BuiltInDocumentProperties("Number of Pages").Value  # 备用方法
            
        doc.Close()
        app.Quit()
        return count if count > 0 else 1
    except Exception as e:
        print(f"文档页数获取失败: {str(e)}")
        return 0
    finally:
        pythoncom.CoUninitialize()


def get_ppt_page_count(file_path):
    """
    针对 XML 格式的 PowerPoint 文件（如 .pptx、.pptm、.ppsm、.potm、.ppsx、.potx）
    使用 python-pptx 获取幻灯片数量
    """
    try:
        return len(Presentation(file_path).slides)
    except Exception as e:
        print(f"PPTX页数获取失败: {str(e)}")
        return 0


def get_ppt_com_page_count(file_path):
    """
    针对传统二进制格式的 PowerPoint 文件（如 .ppt、.pot、.pps、.dpt、.dps、.ett），
    或者当 python-pptx 返回 0 时，采用 COM 自动化方式获取幻灯片数量
    """
    try:
        pythoncom.CoInitialize()
        ppt_app = create_ppt_app()
        # WithWindow=False 防止弹出窗口
        presentation = ppt_app.Presentations.Open(os.path.abspath(file_path), WithWindow=False)
        count = presentation.Slides.Count
        presentation.Close()
        ppt_app.Quit()
        return count if count > 0 else 1
    except Exception as e:
        print(f"PPT页数获取失败: {str(e)}")
        return 0
    finally:
        pythoncom.CoUninitialize()


def get_xls_page_count(file_path):
    """
    遍历 Excel 文件中所有工作表，并求每个工作表的页数之和
    """
    total_pages = 0
    try:
        pythoncom.CoInitialize()
        excel_apps = ['Excel.Application', 'KET.Application', 'Et.Application']
        app = create_office_app(excel_apps)
        
        workbook = app.Workbooks.Open(os.path.abspath(file_path))
        
        for sheet in workbook.Worksheets:
            try:
                pages = sheet.PageSetup.Pages.Count
                if pages < 1:
                    pages = 1
            except Exception:
                pages = 1
            total_pages += pages
        
        workbook.Close()
        app.Quit()
        return total_pages if total_pages > 0 else 1
    except Exception as e:
        print(f"表格页数获取失败: {str(e)}")
        return 1
    finally:
        pythoncom.CoUninitialize()


def get_default_page_count(file_path):
    """对于无法直接统计页数的格式，默认返回1"""
    return 1


def process_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    page_count = 0

    # 定义各类扩展名
    pdf_exts       = ['.pdf']
    word_exts      = ['.doc', '.docx', '.wps', '.wpt', '.dot', '.rtf', '.dotx', '.docm', '.dotm']
    excel_exts     = ['.xls', '.xlt', '.xlsx', '.xlsm', '.xltx', '.xltm', '.xlam', '.xla', '.csv', '.prn', '.dif', '.et']
    pptx_exts      = ['.pptx', '.pptm', '.ppsm', '.potm', '.ppsx', '.potx']
    ppt_legacy_exts = ['.ppt', '.pot', '.pps', '.dpt', '.dps', '.ett']
    html_exts      = ['.xml', '.mht', '.mhtml', '.html', '.htm']
    other_exts     = ['.dbf', '.rtt']
    txt_exts       = ['.txt']

    if ext in pdf_exts:
        page_count = get_pdf_page_count(file_path)
    elif ext in word_exts:
        page_count = get_doc_page_count(file_path)
    elif ext in excel_exts:
        page_count = get_xls_page_count(file_path)
    elif ext in pptx_exts:
        page_count = get_ppt_page_count(file_path)
        # 如果 PPTX 系列返回0，尝试使用 COM 自动化
        if page_count == 0:
            page_count = get_ppt_com_page_count(file_path)
    elif ext in ppt_legacy_exts:
        page_count = get_ppt_com_page_count(file_path)
    elif ext in html_exts or ext in other_exts or ext in txt_exts:
        page_count = get_default_page_count(file_path)
    else:
        page_count = 0  # 不支持的格式返回0

    return page_count


def rename_file(root, filename):
    old_path = os.path.join(root, filename)
    if os.path.isdir(old_path):
        return

    ext = os.path.splitext(filename)[1]
    base_name = os.path.splitext(filename)[0]

    modified_date = get_file_modified_date(old_path)
    page_count = process_file(old_path)

    new_filename = f"{modified_date}-{base_name}-{page_count}{ext}"
    new_path = os.path.join(root, new_filename)

    try:
        os.rename(old_path, new_path)
        print(f"重命名成功: {filename} -> {new_filename}")
    except Exception as e:
        print(f"重命名失败 {filename}: {str(e)}")


def process_folder(folder_path):
    supported_exts = (['.pdf'] +
                      word_exts +
                      excel_exts +
                      pptx_exts +
                      ppt_legacy_exts +
                      html_exts +
                      other_exts +
                      txt_exts)
    for root_dir, dirs, files in os.walk(folder_path):
        for filename in files:
            ext = os.path.splitext(filename)[1].lower()
            if ext in supported_exts:
                rename_file(root_dir, filename)


if __name__ == "__main__":
    # 为方便定义各类型扩展名（供 process_folder 使用）
    word_exts      = ['.doc', '.docx', '.wps', '.wpt', '.dot', '.rtf', '.dotx', '.docm', '.dotm']
    excel_exts     = ['.xls', '.xlt', '.xlsx', '.xlsm', '.xltx', '.xltm', '.xlam', '.xla', '.csv', '.prn', '.dif', '.et']
    pptx_exts      = ['.pptx', '.pptm', '.ppsm', '.potm', '.ppsx', '.potx']
    ppt_legacy_exts = ['.ppt', '.pot', '.pps', '.dpt', '.dps', '.ett']
    html_exts      = ['.xml', '.mht', '.mhtml', '.html', '.htm']
    other_exts     = ['.dbf', '.rtt']
    txt_exts       = ['.txt']

    root = tk.Tk()
    root.withdraw()

    target_folder = filedialog.askdirectory(title="请选择需要处理的文件夹")
    if target_folder:
        process_folder(target_folder)
        messagebox.showinfo("处理完成", "文件夹处理完成！")
    else:
        messagebox.showwarning("未选择文件夹", "没有选择文件夹，程序退出。")
