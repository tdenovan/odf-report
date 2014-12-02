require './lib/odf-report'
require 'faker'

report = ODFReport::Report.new("test/templates/temp_text_field.docx") do |r|

  r.add_text_field("def", 'Faker::Lorem.word')
  r.add_text_field("jkl", 'Faker::Company.catch_phrase')
  r.add_text_field("vwx", 'Faker::Company.duns_number')
  r.add_text_field("table", 'replaced table 1')
  r.add_text_field("table2", 'replaced table 2')

end

report.generate("test/result/test_text_field.docx")
