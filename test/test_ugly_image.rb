require './lib/odf-report'
require 'faker'

report = ODFReport::Report.new("test/templates/temp_image.docx") do |r|

  r.add_image('IMAGE', File.join(Dir.pwd, 'test', 'templates', 'replace.jpeg'))

end

report.generate("test/result/test_ugly_image.docx")
