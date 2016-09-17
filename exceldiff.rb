require 'win32ole'

module Excel
end

def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end
filename = getAbsolutePath("test.xlsx")

excel = WIN32OLE.new('Excel.Application')
WIN32OLE.const_load(excel, Excel)

excel.displayalerts = false

book = excel.Workbooks.Open(filename)
begin
  current_dir = Dir.pwd.gsub("/", "\\")
  book.Worksheets.Item('Sheet2')
  #book.SaveAs('C:\Users\user\Desktop\exceldiff-master\test2.html', :FileFormat => Excel::XlHtml)
  #book.SaveAs('C:\Users\user\Desktop\exceldiff-master\test3.txt', :FileFormat => Excel::XlUnicodeText)
  book.Worksheets.each do |sheet|
    p "Saving sheet:[#{sheet.Name}]"
    sheet.SaveAs("#{current_dir}\\#{sheet.Name}.txt", :FileFormat => Excel::XlUnicodeText)
  end
ensure
  book.Close
  excel.Quit
end
