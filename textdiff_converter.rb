require 'rake'
require 'open3'

FLAG_SAME = 0
FLAG_MODIFIED = 1

HTML_STRING = \
"<html>" \
"<head>" \
"<meta http-equiv=\"Pragma\" content=\"no-cache\">" \
"<meta http-equiv=\"Cache-Control\" content=\"no-cache\">" \
"</head>" \
"<body>" \
"HTML_BODY" \
"</body>" \
"</html>"

NORMAL_TAG = \
"<pre>NORMAL_DATA</pre>"

BEFORE_TAG = \
"<pre><font style=\"background-color:#ff5555\">BEFORE_DATA</font></pre>"

AFTER_TAG = \
"<pre><font style=\"background-color:#00aaff\">AFTER_DATA</font></pre>"

def convert_diff_to_html(base_file, new_file, output_file)
  diff_file = "tmp.diff"
  compare_result = FLAG_SAME

  begin
    Open3.popen3("diff --new-line-format=\"+%L\" --old-line-format=\"-%L\" --unchanged-line-format=\" %L\" #{base_file} #{new_file}") do |i, o, e, w|
      File.open(diff_file, "w") do |file|
        o.each_line do |line|
          #puts line
          file << line
        end
      end
    end

    result_data = ""

    File.open(diff_file) do |file|
      file.each_line do |line|
        if line.start_with?("---") || 
           line.start_with?("+++") || 
           line.start_with?("@@")
          # skip diff command infomation line
        else
          if line =~ /-(.*)/
            result_data += BEFORE_TAG.gsub("BEFORE_DATA", $1)
            compare_result = FLAG_MODIFIED
          elsif line =~ /\+(.*)/
            result_data += AFTER_TAG.gsub("AFTER_DATA", $1)
            compare_result = FLAG_MODIFIED
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

  return compare_result
end

SAME_LINK_TAG = \
"<font style=\"background-color:#ffffff\"><a href=\"LINK\">SAME_HTML</a></font><br>"

MODIFIED_LINK_TAG = \
"<font style=\"background-color:#33ff33\"><a href=\"LINK\">MODIFIED_HTML</a></font><br>"

def create_summary_html(summary, output_file)
  summary_html = ""
  summary.each do |html, compare_result|
    if compare_result == FLAG_SAME
      p "#{html} is same."
      tmp_html      = SAME_LINK_TAG.gsub("LINK", html)
      summary_html += tmp_html.gsub("SAME_HTML", File.basename(html))
    elsif compare_result == FLAG_MODIFIED
      p "#{html} is modified."
      tmp_html      = MODIFIED_LINK_TAG.gsub("LINK", html)
      summary_html += tmp_html.gsub("MODIFIED_HTML", File.basename(html))
    end
  end

  File.open(output_file, "w") do |file|
    p "Create summary html:[#{output_file}]"
    data = HTML_STRING.gsub("HTML_BODY", summary_html)
    file.write(data)
  end
end

if $0 == __FILE__
  base_file = ENV['BASE']
  new_file  = ENV['NEW']
  diff_file = "tmp.diff"
  output_file = "#{File.basename(new_file)}.html"

  convert_diff_to_html(base_file, new_file, output_file)
end

