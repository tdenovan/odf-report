require './lib/odf-report'
require 'faker'

data = {}
series = []
no_series = rand(2..5)
no_columns = rand(2..5)
title = Faker::Lorem.word

no_columns.times do
  column_name = Faker::Lorem.word
  data[column_name] = []
  data[column_name] << rand(1..10) until data[column_name].length == no_series
end

no_series.times do
  series << Faker::Lorem.word
end


report = ODFReport::Report.new("test/templates/temp_columnchart.docx") do |r|

  r.add_chart("TITLE", data, :series => series, :title => title, :type => 'column')

end

report.generate("test/result/test_word_columnchart.docx")
