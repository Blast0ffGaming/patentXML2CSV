# patentXML2CSV.py

from threading import Thread
import csv
from bs4 import BeautifulSoup
import pandas as pd
import os.path
import PySimpleGUI as sg
CONTACT_INFO = "blasty0ff[at]gmail.com"
VESION_NUMBER = "r0.0.2"

# GUI related

# XML parsing related
# import xml.etree.ElementTree as Xet

# Allow background running

# Global state for ending
WindowsClosed = False

# window_FileSelect interface
layout = [
    [
        sg.Text("[01] -> Select a folder:"),
        sg.In(size=(50, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse()
    ],
    [
        sg.Listbox(values=[], enable_events=True,
                   size=(80, 20), key="-FILE LIST-")
    ],
    [
        sg.Text("[02] -> Select a TIPO flavored patent XML file: "),
        sg.Text(size=(40, 1), key="-FILE SELECTED-")
    ],
    [
        sg.Button("Cancel"),
        sg.Button("Convert"),
        sg.Text(size=(60, 1), key="-INFO PROMPT-")
    ],
    [
        sg.ProgressBar(max_value=100, orientation='h',
                       size=(20, 20), key='-PROGRESS CONVERT-', visible=False),
        sg.Text(size=(40, 2), key="-PROGRESS PROMPT-", visible=False),
        sg.Text(size=(60, 2), key="-ERROR PROMPT-", visible=False)
    ]

]


# creating the window_FileSelect
window_FileSelect = sg.Window(
    title="patentXML2CSV (Version: "+VESION_NUMBER+")", layout=layout)
# window_FileSelect=sg.Window(title="hello world", layout=[[]], margins=(100, 50))

# window_FileSelect's event loop
filename = None
while True:

    # get value and event
    event, values = window_FileSelect.read()

    # show list of files to select
    if event == "-FOLDER-":
        folderName = values["-FOLDER-"]
        try:
            # get list of files in folder
            file_list = os.listdir(folderName)
        except:
            file_list = []

        fnames = [
            f for f in file_list if os.path.isfile(os.path.join(folderName, f)) and f.lower().endswith((".xml"))
        ]
        window_FileSelect["-FILE LIST-"].update(fnames)
    elif event == "-FILE LIST-":
        try:
            filename = os.path.join(
                values["-FOLDER-"], values["-FILE LIST-"][0])
            window_FileSelect["-FILE SELECTED-"].update(filename)
        except:
            pass

    # closing the app
    if event == "Cancel" or event == sg.WIN_CLOSED:
        WindowsClosed = True
        break

    if event == "Convert" and filename is None:
        window_FileSelect["-INFO PROMPT-"].update("No file is selected.")

    if event == "Convert" and filename is not None:
        window_FileSelect["-FILE LIST-"].update(disabled=False)
        window_FileSelect["Convert"].update(disabled=True)
        # window_FileSelect["Cancel"].update(disabled=True)
        window_FileSelect["-INFO PROMPT-"].update(
            "This can take a few minutes up to hours depending on file size.")

        window_FileSelect["-ERROR PROMPT-"].update(visible=False)

        def DoXMLParsing():

            WorkingFilename = filename
            # Parsing the XML file
            print("Now reading <"+WorkingFilename+">")

            window_FileSelect["-PROGRESS CONVERT-"].update(
                visible=True)
            window_FileSelect["-PROGRESS PROMPT-"].update(
                visible=True)

            window_FileSelect["-PROGRESS CONVERT-"].UpdateBar(
                0)
            window_FileSelect["-PROGRESS PROMPT-"].update("Initalizing...")

            # 創建Beautiful Soup對象並解析XML數據
            with open(WorkingFilename, 'r', encoding='utf-8') as xml_file:
                soup = BeautifulSoup(xml_file, 'lxml')

            total_path_count = 0
            total_patent_count = 0

            # 找到所有不同的標籤路徑，這將成為CSV的列標題
            unique_paths = set()
            for patent in soup.find_all('tw-patent-grant'):
                paths = set('/'.join([ancestor.name for ancestor in tag.parents] + [tag.name])
                            for tag in patent.find_all(recursive=True))
                unique_paths.update(paths)

                # End early if already closed
                if WindowsClosed:
                    return

                # 印出每個路徑的額外信息
                # for path in paths:
                #     print("Locating <" + repr(path) + ">")

                total_path_count += len(paths)
                total_patent_count += 1
                # print("Currently <"+str(total_path_count)+"> found")
                window_FileSelect["-PROGRESS PROMPT-"].update(
                    "Finding <"+str(total_path_count)+"> nodes in <"+str(total_patent_count)+"> paths")

            print("Total of <"+str(total_path_count)+"> and <" +
                  str(total_patent_count)+"> paths found")

            if unique_paths == set():
                WindowsError = "File Format Error"
                print("No tw-patent-grant found.")

                window_FileSelect["-PROGRESS CONVERT-"].update(
                    visible=False)
                window_FileSelect["-PROGRESS PROMPT-"].update(
                    visible=False)
                window_FileSelect["-ERROR PROMPT-"].update(
                    "The file might not be a TIPO flavored patent XML file, or that TIPO has updated it's XML standard. In the latter case, please report it to <" +
                    CONTACT_INFO+"> for this. ")
                window_FileSelect["-ERROR PROMPT-"].update(visible=True)

                window_FileSelect["Convert"].update(disabled=False)
                # window_FileSelect["Cancel"].update(disabled=True)

                window_FileSelect["-INFO PROMPT-"].update(
                    "The converion has been terminated.")

                return

            # 創建CSV文件並打開以寫入模式，使用 'utf-8-sig' 編碼
            with open('output.csv', 'w', newline='', encoding='utf-8-sig') as csv_file:
                writer = csv.writer(csv_file)

                # 寫入CSV文件的標題行
                writer.writerow(list(unique_paths))

                # 遍歷每個 <tw-patent-grant> 元素
                counter_i = 0
                total_patent_count = len(soup.find_all('tw-patent-grant'))
                counter_p = 0
                for patent in soup.find_all('tw-patent-grant'):

                    counter_p += 1

                    # 初始化一個空的字典，用於存儲每個標籤的內容
                    data = {}

                    # 遍歷所有不同的標籤路徑
                    for tag_path in unique_paths:

                        # 找到具有該標籤路徑的元素，如果存在，將其內容添加到字典中
                        elements = patent.find_all(
                            name=tag_path.split('/')[-1], recursive=True)
                        if elements:
                            # 將多個內容以分號分隔符合併為一個字符串
                            data[tag_path] = '; '.join(element.text.strip()
                                                       for element in elements)
                        else:
                            data[tag_path] = ''

                        # End early if already closed
                        if WindowsClosed:
                            return

                        # print("Traversing <" + repr(tag_path) + ">")
                        counter_i = counter_i+1

                    window_FileSelect["-PROGRESS CONVERT-"].UpdateBar(
                        counter_p / total_patent_count * 100)
                    window_FileSelect["-PROGRESS PROMPT-"].update(
                        '{}% ({})'.format(int(counter_p / total_patent_count * 100), str(counter_p)+"/"+str(total_patent_count)))

                    # 將字典的值轉換為一個CSV行並寫入CSV文件
                    writer.writerow([data[tag_path]
                                    for tag_path in unique_paths])

                    # print("Total of <"+str(counter_i)+"> items traversed")

            window_FileSelect["-FILE LIST-"].update(disabled=False)
            window_FileSelect["Convert"].update(disabled=False)
            # window_FileSelect["Cancel"].update(disabled=True)

            window_FileSelect["-INFO PROMPT-"].update(
                "The converion has completed.")

        Thread(target=DoXMLParsing).start()

window_FileSelect.close()
