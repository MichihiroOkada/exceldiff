#require 'exceldiff.rb'
require 'rake'

HTML_STRING = \
"<html>" \
"<body>" \
"HTML_BODY" \
"</body>" \
"</html>"

NORMAL_TAG = \
"NORMAL_DATA<br>"

BEFORE_TAG = \
"<font style=\"background-color:#ff5555\">BEFORE_DATA</font><br>"

AFTER_TAG = \
"<font style=\"background-color:#00aaff\">AFTER_DATA</font><br>"

def convert_diff_to_html(base_file, new_file, output_file)
  diff_file = "tmp.diff"

  begin
    sh "diff -u #{base_file} #{new_file} > #{diff_file}" do |ok, status|
    end

    result_data = ""

    File.open(diff_file) do |file|
      file.each_line do |line|
        if line.start_with?("---") || 
           line.start_with?("+++") || 
           line.start_with?("@@")
          # skip diff command infomation
        else
          if line =~ /-(.*)/
            result_data += BEFORE_TAG.gsub("BEFORE_DATA", $1)
            p result_data
          elsif line =~ /\+(.*)/
            result_data += AFTER_TAG.gsub("AFTER_DATA", $1)
            p result_data
          else
            result_data += NORMAL_TAG.gsub("NORMAL_DATA", line)
          end
        end
      end # end of file.each_line
    end # end of File.open

    File.open(output_file, "w") do |file|
      data = HTML_STRING.gsub("HTML_BODY", result_data)
      file.write(data)
    end
  ensure
    rm_f diff_file
  end
end

#task :convert_diff_to_html do
task :convert do
  base_file = ENV['BASE']
  new_file  = ENV['NEW']
  diff_file = "tmp.diff"
  output_file = "#{File.basename(new_file)}.html"

  convert_diff_to_html(base_file, new_file, output_file)
end

