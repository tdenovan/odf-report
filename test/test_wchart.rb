require './lib/odf-report'
require 'faker'

alphabet = {
  'Start of Semester' => rand(-20..20),
  'Jan' => rand(-20..20),
  'Feb' => rand(-20..20),
  'Mar' => rand(-20..20),
  'Apr' => rand(-20..20),
  'May' => rand(-20..20),
  'Jun' => rand(-20..20),
  'End Of Semester' => 0
}

report = ODFReport::Report.new("test/templates/temp_wchart.docx") do |r|

  r.add_chart("TITLE", alphabet)

end

report.generate("test/result/test_word_wchart.docx")