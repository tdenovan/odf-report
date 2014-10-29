require './lib/odf-report'
require 'faker'

hash = {}

rand(4..10).times do
  hash[Faker::Lorem.word] = rand(1..20)
end

report = ODFReport::Report.new("test/templates/temp_psheet.xlsx") do |r|

  r.add_chart("CHART_01", hash, :series => 'Hello', :title => 'Goodbye')

end

report.generate("test/result/test_psheet.xlsx")
