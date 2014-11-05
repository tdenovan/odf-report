require './lib/odf-report'
require 'faker'

data = {}
series = []
colors = []
no_series = rand(5..10)
no_columns = rand(5..10)
title = Faker::Lorem.word

no_columns.times do
  column_name = Faker::Lorem.word
  data[column_name] = []
  data[column_name] << rand(1..10) until data[column_name].length == no_series
end

no_series.times do
  series << Faker::Lorem.word
  colors << "#{rand(1..6)}.#{rand(1...6)}".to_f
end


report = ODFReport::Report.new("test/templates/temp_linechart.docx") do |r|

  r.add_chart("chart", data, :series => series, :title => title, :type => 'line', :colors => colors)

end

report.generate("test/result/test_word_linechart.docx")
