require './lib/odf-report'
require 'faker'

alphabet = {
  'Start' => 30,
  'Q1' => rand(-10..20),
  'Q2' => rand(-10..20),
  'Q3' => rand(-10..20),
  'Q4' => rand(-10..20)
}

report = ODFReport::Report.new("test/templates/temp_wchart.docx") do |r|

  r.add_chart("TITLE", alphabet)

end

report.generate("test/result/test_word_wchart.docx")
