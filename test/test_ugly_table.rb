require './lib/odf-report'
require 'ostruct'
require 'faker'
require 'launchy'

@col1 = []

10.times do |i|
  @col1 << {
    :name => Faker::Name.name,
    :id => i,
    :city => Faker::Address.city,
    :address => Faker::Address.street_address
  }
end

report = ODFReport::Report.new("test/templates/temp_table.docx") do |r|

  r.add_field("HEAD_01", 'ID')
  r.add_field("HEAD_02", 'Name')
  r.add_field("HEAD_03", 'Address')

  r.add_table("TABLE_01", @col1, :header=>true) do |t|
    t.add_column(:field_01, :id)
    t.add_column(:field_02, :name)
    t.add_column(:field_03, :address)
  end

end

report.generate("test/result/test_ugly_table.docx")
