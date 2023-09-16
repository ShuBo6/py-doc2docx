import os
from win32com import client as wc


def save_doc_to_docx(rawpath):
    '''
    :param rawpath: 传入和传出文件夹的路径
    :return: None
    '''

    try:
        word = wc.Dispatch("Word.Application")
    except Exception as e:
        print("无法启动 Word 应用程序：")
        print(e)
        return

    try:
        filenamelist = os.listdir(rawpath)
    except Exception as e:
        print("无法列出目录：")
        print(e)
        word.Quit()
        return

    for i in os.listdir(rawpath):
        if i.endswith('.doc') and not i.startswith('~$'):
            print("正在处理文件：", i)
            try:
                doc = word.Documents.Open(rawpath + i)
                rename = os.path.splitext(i)
                doc.SaveAs(path + rename[0] + '.docx', 12)
                doc.Close()
            except Exception as e:
                print("处理文件时发生错误：", i)
                print(e)

    word.Quit()


if __name__ == '__main__':
    path = 'C:\\Users\\81418\\Desktop\\doc2docx\\xxxx\\'
    save_doc_to_docx(path)
