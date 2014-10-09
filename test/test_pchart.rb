require './lib/odf-report'
require 'faker'

alphabet = {
  'Alpha' => rand(10),
  'Beta' => rand(10),
  'Gamma' => rand(10),
  'Delta' => rand(10)
}

report = ODFReport::Report.new("test/templates/temp_piechart.docx") do |r|

  r.add_title("TITLE", "New Chart Name")
  r.add_series("QUANTITY", "New Quantity")
  r.add_chart("CHART_01", alphabet)

end

report.generate("test/result/test_word_piechart.docx")
