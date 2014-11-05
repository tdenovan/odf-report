require './lib/odf-report'
require 'faker'

hash = {}
colors = []

6.times do
  hash[Faker::Lorem.word] = rand(1..20)
  colors << "#{rand(1..6)}.#{rand(1...6)}".to_f
end



report = ODFReport::Report.new("test/templates/temp_donutchart.docx") do |r|

  r.add_chart("CHART_01", hash, :series => 'Hello', :type => 'doughnut', :colors => colors, :labels => :enabled)

end

report.generate("test/result/test_word_donutchart.docx")
