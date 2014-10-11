require './lib/odf-report'
require 'faker'

report = ODFReport::Report.new("test/templates/temp_image.docx") do |r|

  r.add_image('IMAGE', File.join(Dir.pwd, 'test', 'templates', 'replace.jpeg'))
  r.add_image('IMAGE2', File.join(Dir.pwd, 'test', 'templates', 'copy.jpeg'))

end

report.generate("test/result/test_word_image.docx")
