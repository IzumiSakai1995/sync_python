
```python
import win32com.client as win32

# 打开 Word 应用程序和文档
word = win32.Dispatch('Word.Application')
doc = word.Documents.Open(r'D:\test.docx')

# 更新自定义属性 version
doc.CustomDocumentProperties('version').Value = 'V100R021C10SPC101'

# 刷新域
doc.Fields.Update()

# 保存并关闭文档
doc.Save()
doc.Close()

# 退出 Word 应用程序
word.Quit()

```

import win32com.client as win32

# 打开 Word 应用程序和文档
word = win32.Dispatch('Word.Application')
doc = word.Documents.Open(r'D:\test.docx')

# 更新自定义属性 version
doc.CustomDocumentProperties('version').Value = 'V100R021C10SPC101'

# 刷新域
doc.Fields.Update()

# 保存并关闭文档
doc.Save()
doc.Close()

# 退出 Word 应用程序
word.Quit()

import win32com.client as win32

# 打开 Word 应用程序和文档
word = win32.Dispatch('Word.Application')
doc = word.Documents.Open(r'D:\test.docx')

# 更新自定义属性 version
doc.CustomDocumentProperties('version').Value = 'V100R021C10SPC101'

# 刷新域
doc.Fields.Update()

# 保存并关闭文档
doc.Save()
doc.Close()

# 退出 Word 应用程序
word.Quit()
