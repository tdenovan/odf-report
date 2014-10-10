require './lib/odf-report'
require 'faker'

alphabet = {
  'Start' => 20,
  'Jan' => 10,
  'Feb' => -20,
  'Mar' => -20,
  'Apr' => -10,
  'May' => 30,
  'Jun' => 20
}

report = ODFReport::Report.new("test/templates/temp_wchart.docx") do |r|

  r.add_chart("TITLE", alphabet)

end

report.generate("test/result/test_word_wchart.docx")