require './lib/odf-report'
require 'faker'

alphabet = {
  'Alpha' => rand(1..10),
  'Beta' => rand(1..10),
  'Gamma' => rand(1..10),
  'Delta' => rand(1..10)
}

things = {
  'Apple' => rand(1..10),
  'Banana' => rand(1..10),
  'Cabbage' => rand(1..10),
  'Doggy' => rand(1..10)
}

@col1 = []
10.times do |i|
  @col1 << {
    :name => Faker::Name.name,
    :id => i,
    :city => Faker::Address.city,
    :address => Faker::Address.street_address
  }
end

@col2 = []
10.times do |i|
  @col2 << {
    :name => Faker::Name.name,
    :id => i,
    :city => Faker::Address.city,
  }
end

report = ODFReport::Report.new("test/templates/temp_all.docx") do |r|

  # Text
  r.add_field("CLIENT_NAME", Faker::Company.name)
  r.add_field("DATE_TODAY", Time.new(2002, 10, 31, 2, 2, 2, "+02:00"))
  r.add_field("TEXT_01", Faker::Company.catch_phrase)
  r.add_field("TEXT_02", Faker::Company.catch_phrase)

  # Chart
  r.add_field("CHART_01", "New Chart Name")
  r.add_chart("CHART_01", alphabet)
  r.add_field("SOME TITLE", "HORRAY")
  r.add_chart("CHART_02", things)

  # Image
  # r.add_image('IMAGE_01', File.join(Dir.pwd, 'test', 'templates', 'replace.jpeg'))
  # r.add_image('IMAGE_02', File.join(Dir.pwd, 'test', 'templates', 'copy.jpeg'))

  # Table
  r.add_field("HEAD_01", 'ID1')
  r.add_field("HEAD_02", 'Name1')
  r.add_field("HEAD_03", 'Address1')
  r.add_field("HEAD_04", 'City1')

  r.add_table("TABLE_01", @col1) do |t|
    t.add_column(:field_01, :id)
    t.add_column(:field_02, :name)
    t.add_column(:field_03, :address)
    t.add_column(:field_04, :city)
  end

  r.add_field("HEAD_11", 'ID2')
  r.add_field("HEAD_12", 'Name2')
  r.add_field("HEAD_13", 'Address2')

  r.add_table("TABLE_02", @col2) do |t|
    t.add_column(:field_11, :id)
    t.add_column(:field_12, :name)
    t.add_column(:field_13, :city)
  end

end

report.generate("test/result/test_word_all.docx")
