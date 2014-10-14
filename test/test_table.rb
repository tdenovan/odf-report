require './lib/odf-report'
require 'ostruct'
require 'faker'
require 'launchy'

@col1 = []

3.times do |i|
  image = i < 2 ? File.join(Dir.pwd, 'test', 'templates', 'image_01.jpg') : File.join(Dir.pwd, 'test', 'templates', 'image_02.jpg')
  @col1 << {
    :name => Faker::Name.name,
    :image => image,
    :city => Faker::Address.city,
    :address => Faker::Address.street_address
  }
end

report = ODFReport::Report.new("test/templates/temp_table.docx") do |r|

  r.add_field("HEAD_01", 'ID')
  r.add_field("HEAD_02", 'Name')
  r.add_field("HEAD_03", 'Address')
  r.add_field("HEAD_04", 'City')

  r.add_table("TABLE_01", @col1, :header=>true) do |t|
    t.add_image(:table_image, :image)
    t.add_column(:field_02, :name)
    t.add_column(:field_03, :address)
    t.add_column(:field_04, :city)
  end

  # r.add_table("TABLE_02", @col2) do |t|
  #   t.add_column(:field_04, :idx)
  #   t.add_column(:field_05, :name)
  #   t.add_column(:field_06, 'address')
  #   t.add_column(:field_07, :phone)
  #   t.add_column(:field_08, :zip)
  # end

  # r.add_table("TABLE_03", @col3, :header=>true) do |t|
  #   t.add_column(:field_01, :idx)
  #   t.add_column(:field_02, :name)
  #   t.add_column(:field_03, :address)
  # end

  # r.add_table("TABLE_04", @col4, :header=>true, :skip_if_empty => true) do |t|
  #   t.add_column(:field_01, :idx)
  #   t.add_column(:field_02, :name)
  #   t.add_column(:field_03, :address)
  # end

  # r.add_table("TABLE_05", @col5) do |t|
  #   t.add_column(:field_01, :idx)
  #   t.add_column(:field_02, :name)
  #   t.add_column(:field_03, :address)
  # end

  # r.add_image("graphics1", File.join(Dir.pwd, 'test', 'templates', 'piriapolis.jpg'))
  # r.add_image("graphics2", File.join(Dir.pwd, 'test', 'templates', 'rails.png'))

end

report.generate("test/result/test_word_table.docx")
