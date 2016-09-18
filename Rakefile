# coding: sjis

require './textdiff_converter.rb'
require 'win32ole'
require 'kconv'
require 'rake'

module Excel
end

$excel = WIN32OLE.new('Excel.Application')
WIN32OLE.const_load($excel, Excel)

# 警告無効化
$excel.displayalerts = false

def getAbsolutePath(filename)
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

def excel2text(filename, style)
  filename = getAbsolutePath(filename)
  p filename

  book = $excel.Workbooks.Open(filename)
  begin
    output_dir = File.dirname(filename)
    #book.SaveAs("#{output_dir}\\test2.html", :FileFormat => Excel::XlHtml)
    #book.SaveAs("#{output_dir}\\test3.txt", :FileFormat => Excel::XlUnicodeText)
    book.Worksheets.each do |sheet|
      p "Saving sheet:[#{sheet.Name}]"
      p "Output file:[#{output_dir}\\#{sheet.Name}.txt]"
      sheet.SaveAs("#{output_dir}\\#{sheet.Name}.txt", :FileFormat => style)
    end
  ensure
    book.Close
    $excel.Quit
  end
end

task :excel2text do
  base_dir = ENV['BASE_DIR']
  new_dir  = ENV['NEW_DIR']
  
  Dir.glob("#{base_dir}/**/*.{xls,xlsx,xlsm}") do |file|
    excel2text(file, Excel::XlUnicodeText)
  end
  Dir.glob("#{base_dir}/**/*.txt") do |file|
    #s = File.open(file, :encoding => Encoding::UTF_8).read()
    #File.open(file, "w").write(s.tosjis)
    sh "nkf -s --overwrite #{file}"
  end
  
  Dir.glob("#{new_dir}/**/*.{xls,xlsx,xlsm}") do |file|
    excel2text(file, Excel::XlUnicodeText)
  end
  Dir.glob("#{new_dir}/**/*.txt") do |file|
    #s = File.open(file, :encoding => Encoding::UTF_8).read()
    #File.open(file, "w").write(s.tosjis)
    sh "nkf -s --overwrite #{file}"
  end
end  

task :excel2text_with_style do
  base_dir = ENV['BASE_DIR']
  new_dir  = ENV['NEW_DIR']
  
  Dir.glob("#{base_dir}/**/*.{xls,xlsx,xlsm}") do |file|
    excel2text(file, Excel::XlTextPrinter)
  end

  Dir.glob("#{new_dir}/**/*.{xls,xlsx,xlsm}") do |file|
    excel2text(file, Excel::XlTextPrinter)
  end
end  

#task :convert_diff_to_html do
task :textdiff2html do
  base_dir = ENV['BASE_DIR']
  new_dir  = ENV['NEW_DIR']
  output_dir  = ENV['OUTPUT_DIR']
  diff_file = "tmp.diff"

  summary = {}

  Dir.glob("#{new_dir}/**/*.txt") do |file|
    base_file = file.gsub(new_dir, base_dir)
    output_file = "#{output_dir}/#{File.basename(file, '.*')}.html"

    p "Output file:[#{output_file}]"
    if File.exists?(base_file) 
      compare_result = convert_diff_to_html(base_file, file, output_file)
      summary[output_file] = compare_result
    else
      p "!!new sheet detected!!"
    end
  end

  create_summary_html(summary, "index.html")
end

