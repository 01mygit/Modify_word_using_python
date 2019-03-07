import win32com
from win32com.client import Dispatch
import glob

w = win32com.client.Dispatch('kwps.Application')
w.Visible = 0
w.DisplayAlerts = 0

def process():
    filenames = glob.glob(root + '\*.doc')
    for filename in filenames:
        print(filename)
        doc = w.Documents.Open(FileName=filename)

        # # 文档最开始插入文字
        # insert = filename.split('.')[0] + '\n'
        # myRange = doc.Range(0, 0)
        # myRange.InsertBefore(insert)

        par = doc.Range(10, doc.Content.End)
        par.ParagraphFormat.LineSpacing = 12

        w.ActiveDocument.Select()
        w.Selection.Font.Name = "微软雅黑"
        w.Selection.Font.Size = "12"
        # 删除空行，这里数量是1，因为回车占一个字符
        for each in w.ActiveDocument.Paragraphs:
            if each.Range.Words.Count == 1:
                each.Range.Delete()
        print("已处理：" + filename)
        # 保存为PDF
        pdf_name = filename.split('.')[0]
        doc.SaveAs(pdf_name, FileFormat=17)
        doc.Close()
    print("处理完毕！")


if __name__ =='__main__':
    root = r'F:\github\Modify_word_using_python\word'
    process()


