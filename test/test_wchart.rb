require './lib/odf-report'
require 'faker'

alphabet = {
  'Start of Semester' => rand(0..20),
  'Jan' => rand(0..20),
  'Feb' => rand(0..20),
  'End Of Semester' => 0
}

report = ODFReport::Report.new("test/templates/temp_wchart.docx") do |r|

  r.add_chart("TITLE", alphabet, :type => 'waterfall')

end

report.generate("test/result/test_word_wchart.docx")