require './lib/odf-report'
require 'faker'

alphabet = {
  'Alpha' => 1,
  'Beta' => 2,
  'Gamma' => 3,
  'Delta' => 4
}

things = {
  'Apple' => 5,
  'Banana' => 6,
  'Cabbage' => 7,
  'Doggy' => 8
}

bar = {
  'Elephant' => [9, 10, 11],
  'Fudge' => [12, 13, 14],
  'Google' => [15, 16, 17],
  'Hippo' => [18, 19, 20]
}

column = {
  'Io' => [21, 22, 23],
  'Jiggly' => [24, 25, 26],
  'Kinky' => [27, 28, 29],
  'Loop' => [30, 31, 32]
}

more_things = {
  'Apple' => 32,
  'Banana' => 33,
  'Cabbage' => 34,
  'Doggy' => 35
}

colors = [
  ["#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f],
  ["#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f],
  ["#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f],
  ["#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f, "#{rand(1..6)}.#{rand(1..5)}".to_f],
  ["#{rand(1..6)}.#{rand(1..5)}".to_f]
]


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
  # r.add_chart("CHART_01", alphabet, :series => 'Pi', :type => 'pie', :colors => colors[0])

  # r.add_chart("CHART_02", things, :series => 'Mmm... Doughnut...', :type => 'doughnut', :colors => colors[1])

  # r.add_chart("CHART_03", column, :series => ['Abc', 'Def', 'Ghi'], :type => 'column', :colors => colors[2])

  # r.add_chart("CHART_04", bar, :series => ['Jkl'],  :type => 'bar', :colors => colors[3])

  # r.add_chart('chart_05', more_things,  :type => 'waterfall', :colors => colors[4])

  r.add_chart("CHART_01", alphabet, :series => 'Pi', :type => 'pie', :colors => colors[0], :labels => true, :legend => false)

  r.add_chart("CHART_02", things, :title => 'Mmm... Doughnut...', :type => 'doughnut', :colors => colors[1], :labels => true, :legend => false)

  r.add_chart("CHART_03", column, :series => ['Abc', 'Def', 'Ghi'], :title => 'Columbus', :type => 'column', :colors => colors[2], :labels => true, :legend => false)

  r.add_chart("CHART_04", bar, :series => ['Jkl'], :title => 'Barhopping Chart', :type => 'bar', :colors => colors[3], :labels => true, :legend => false)

  r.add_chart('chart_05', more_things, :title => 'Don\'t Go Chasing Waterfalls', :type => 'waterfall', :colors => colors[4], :labels => true, :legend => false)

  r.add_chart("chart_06", bar, :series => ['Jkl', 'Mno', 'Pqr'], :title => 'Barhopping Chart', :type => 'line', :colors => colors[3], :labels => true, :legend => false)

  # Image
  # r.add_image('IMAGE_01', File.join(Dir.pwd, 'test', 'templates', 'replace.jpeg'))
  # r.add_image('IMAGE_02', File.join(Dir.pwd, 'test', 'templates', 'copy.jpeg'))

  # Table
  r.add_field("HEAD_01", 'ID1')
  r.add_field("HEAD_02", 'Name1')
  r.add_field("HEAD_03", 'Address1')
  # r.add_field("HEAD_04", 'City1')

  r.add_table("TABLE_01", @col1) do |t|
    t.add_column(:field_01, :id)
    t.add_column(:field_02, :name)
    t.add_column(:field_03, :address)
    # t.add_column(:field_04, :city)
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
