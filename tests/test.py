import win32com.client

word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Open('C:\\Users\\Altair\\Desktop\\转博工作报告-新.md.html', False)
doc.SaveAs('C:\\Users\\Altair\\Desktop\\test.doc', FileFormat=0)
doc.Close()
word.Quit()