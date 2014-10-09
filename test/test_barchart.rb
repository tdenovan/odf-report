require './lib/odf-report'
require 'faker'

alphabet = {
  'Alpha' => [1, 2, 3],
  'Beta' => [4, 5, 6],
  'Gamma' => [7, 8, 9],
  'Delta' => [0, 1, 2]
}

report = ODFReport::Report.new("test/templates/temp_barchart.docx") do |r|

  r.add_series("SERIES_01", "Replaced Series")
  r.add_series("SERIES_02", "Another Series")
  r.add_series("SERIES_03", "Substituted Series")
  r.add_chart("TITLE", alphabet)

end

report.generate("test/result/test_word_barchart.docx")
