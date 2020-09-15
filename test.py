import xlwings as xw
app=xw.App(visible=True,add_book=False)
wb=app.books.open('工装台账汇总.xlsx')
print(wb.fullname)
