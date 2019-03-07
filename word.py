import os
import win32com
from win32com.client import Dispatch

w = win32com.client.Dispatch('kwps.Application')
w.Visible = 0
w.DisplayAlerts = 0

def process():
    for root, dirnames, filenames in os.walk(r'C:\Users\page\Desktop\Morvan\word'):
        for filename in filenames:
            print(filename)
            classRTF = os.path.join(root, filename)
            doc = w.Documents.Open(FileName=classRTF)
            # # 文档最开始插入文字
            # insert = filename.split('.')[0] + '\n'
            # myRange = doc.Range(0, 0)
            # myRange.InsertBefore(insert)

            par = doc.Range(10, doc.Content.End)
            par.ParagraphFormat.LineSpacing = 12

            w.ActiveDocument.Select()
            w.Selection.Font.Name = "微软雅黑"
            w.Selection.Font.Size = "12"
            # 3.删除空行，这里数量是1，因为回车占一个字符
            for each in w.ActiveDocument.Paragraphs:
                if each.Range.Words.Count == 1:
                    each.Range.Delete()
            print("已处理：" + classRTF)
            # 保存为PDF
            pdf_name = filename.split('.')[0]
            pdf_name = os.path.join(root, pdf_name)
            doc.SaveAs(pdf_name, FileFormat=17)
            doc.Close()
            exit(0)
    print("处理完毕！")


if __name__ =='__main__':
    process()


