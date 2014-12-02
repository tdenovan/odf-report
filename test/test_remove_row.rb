require './lib/odf-report'
require 'faker'

report = ODFReport::Report.new("test/templates/temp_remove_row.docx") do |r|
  r.add_variables('qwerty', {apple: 3, banana: 2})
  r.add_variables('cabbage', 5})
  r.add_variables('dog', nil)
  r.add_variables('asdfgh', {elephant: nil})
end

report.generate("test/result/test_remove_row.docx")
