require './lib/odf-report'
require 'faker'

data = {}
series = []
no_series = rand(2..5)
# no_series = 3
no_columns = rand(5..10)
# no_columns = 6
title = Faker::Lorem.word

no_columns.times do
  column_name = Faker::Lorem.word
  data[column_name] = []
  data[column_name] << rand(1..10) until data[column_name].length == no_series
end

no_series.times do
  series << Faker::Lorem.word
end


report = ODFReport::Report.new("test/templates/temp_csheet.xlsx") do |r|

  r.add_chart("TITLE", data, :series => series, :title => title)

end

report.generate("test/result/test_csheet.xlsx")
