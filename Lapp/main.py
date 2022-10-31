from Saver.Excel import ExcelProcessor

ex = ExcelProcessor()
ex.setAllRequired(['students.xlsx'], 'dddddd.xlsx', [], 0)

ex.setconditions(['yy@gmail.com'])
ex.setColumnKeys(['H',7])

if ex.is_Error() is False:
    ex.readingAllFiles()
    ex.readActualData()
    ex.filter()
if ex.is_Error():
    print("Sorry we have encountered some error ERROR:\n",ex.getError())
else:
    print("Finished processed check your file\n",ex.savedFileName())
