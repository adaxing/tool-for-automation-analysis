# tool-for-automation-analysis

### old&new 
- .doc is for binary format pre 2007, .docx is for xml format new, which has more features
- openpyxl
  - workbook.get_sheet_by_name(sheetname) -> workbook[sheetname]
  - workbook.remove_sheet(sheetname) -> del workbook[sheetname] or workbook.remove(sheetname)
  
### tips
- load & copy & paste & save files
  - os 可以创建移动重命名删除 可一层一层读取 需循环读取在文档中所有文件 shutil可以文件复制移动（替代os
  - os.walk 会走到根 返回(cur_dir, sub_dir, filename) 所以想要只走一个level 那么可以用分离separator '\\' 这样抉择哪一个level
  - pathlib.Path 会控制全部文件 和glob配合可以在具体dir搜索特定文件 如Path('directory_1').glob('**/*.txt') 
  - glob 可以对特定文件进行搜索获取 glob('**/*.[txt/xlsx/py..]')
  - word
    - pass data to word need to be STRING type
  - pdf
    - Pdf_load 可能会出现encrypt的问题 需要pdf.decrypt() 
  - markdown file .md
    - 处理数据时 需考虑语法中的符号换行符
  - excel
    - “Strict Open XML Spreadsheet (xlsx)” and “Excel Workbook (xlsx)” have same extension, only “Excel Workbook (xlsx)” is supported by openpyxl library
    - copy&paste
      - copy worksheet within same workbook: 
      ``` 
          source = wb.load_workbook(path)
          target = wb.copy_worksheet(source)
          target.title = 'a'  
      ```
      - copy, paste worksheet with merged cell, will raise error: unexpected keyword argument 'min_col' 
      - 但是不能跨workbook复制粘贴worksheet(文档上有写 也可尝试跨workbook复制粘贴 会出现文件有损 最好的方法是iter_rows() 获取每个row 再循环每个cell 获取cell.value到list 
      这样就可以用dataframe.to_excel() 
      - dataframe.to_excel()会覆盖原有的文件 只留保存的sheet 如果只想对部分更改 可用pd.ExcelWriter
      - pandas无视样式 openpyxl会保存原有的样式 
- 样式
  - Openpyxl 可以调整样式-> 对齐，居中，线形， 边框，字体 
  ```
    from openpyxl.styles import Alignment
    from openpyxl.styles import Side, Border
    from openpyxl.styles import Font
  ```
  - pandas调整样式->DataFrame.style 返回Styler object 需要给callback func(conditional formating)
   

- 语法 
  - strptime: 解析字符串中蕴含的时间; strftime: 转换成所想的格式
  - generator 函数 每次被call生成可iterate set 
  - & 对比两个条件是 需要对比相同的data type
  - 将最后一个el换到开头->list.insert(0, list.pop())
  - ord()将字母转换数字 unicode char to integer; chr() 将数字转换字母 integer to unicode char 
    - A-Z = 65-90

  
  
