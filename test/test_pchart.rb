require './lib/odf-report'
require 'faker'

hash = {}

6.times do
  hash[Faker::Lorem.word] = rand(1..20)
end

colors = [1.0, 1.1, 1.2, 1.3, 1.4, 1.5]

report = ODFReport::Report.new("test/templates/temp_piechart.docx") do |r|

  r.add_chart("CHART_01", hash, :series => 'Hello', :title => 'Goodbye', :type => 'pie', :colors => colors)

end

report.generate("test/result/test_word_piechart.docx")
