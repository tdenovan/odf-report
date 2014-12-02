require './lib/odf-report'
require 'faker'

report = ODFReport::Report.new("test/templates/temp_remove_row.docx") do |r|
  r.add_field 'apple', 3
  r.add_field 'banana', 2
  r.add_field 'cabbage', 5
  r.add_field 'doggy', nil
  r.add_field 'elephant', nil
end

report.generate("test/result/test_remove_row.docx")
