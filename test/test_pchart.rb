require './lib/odf-report'
require 'faker'

hash = {}

rand(4..10).times do
  hash[Faker::Lorem.word] = rand(1..20)
end

report = ODFReport::Report.new("test/templates/temp_piechart.docx") do |r|

  r.add_title("TITLE", "New Chart Name")
  r.add_series("QUANTITY", "New Quantity")
  r.add_chart("CHART_01", hash)

end

report.generate("test/result/test_word_piechart.docx")
