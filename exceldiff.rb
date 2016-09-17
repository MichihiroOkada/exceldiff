require 'win32ole'

module Excel
end

def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end
filename = getAbsolutePath("sample1.excels")

excel = WIN32OLE.new('Excel.Application')
WIN32OLE.const_load(excel)

book = excel.Workbooks.Open(filename)
begin
  book.Worksheets.each do |sheet|
    sheet.UsedRange.Rows.each do |row|
      record = []
      row.Columns.each do |cell|
        record << cell.Value
      end
      puts record.join(",")
    end
  end
ensure
  book.Close
  excel.Quit
end
