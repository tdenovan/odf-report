require './lib/odf-report'
require 'faker'

data = {}
no_series = rand(2..5)
no_columns = rand(2..5)

no_columns.times do
  column_name = Faker::Lorem.word
  data[column_name] = []
  data[column_name] << rand(1..10) until data[column_name].length == no_series
end


report = ODFReport::Report.new("test/templates/temp_columnchart.docx") do |r|

  r.add_series("SERIES_01", "Replaced Series")
  r.add_series("SERIES_02", "Another Series")
  r.add_series("SERIES_03", "Substituted Series")
  r.add_chart("TITLE", data)

end

report.generate("test/result/test_word_columnchart.docx")
