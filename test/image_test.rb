require './lib/odf-report'
require 'faker'


  class Item
    attr_accessor :name, :inner_text
    def initialize(_name,  _text)
      @name=_name
      @inner_text=_text
    end
  end



    @items = []
    3.times do

      text = <<-HTML
        <p>#{Faker::Lorem.sentence} <em>#{Faker::Lorem.sentence}</em> #{Faker::Lorem.sentence}</p>
        <p>#{Faker::Lorem.sentence} <strong>#{Faker::Lorem.paragraph(3)}</strong> #{Faker::Lorem.sentence}</p>
        <p>#{Faker::Lorem.paragraph}</p>
        <blockquote>
          <p>#{Faker::Lorem.paragraph}</p>
          <p>#{Faker::Lorem.paragraph}</p>
        </blockquote>
        <p style="margin: 150px">#{Faker::Lorem.paragraph}</p>
        <p>#{Faker::Lorem.paragraph}</p>
      HTML

      @items << Item.new(Faker::Name.name, text)

    end


    item = @items.pop

    report = ODFReport::Report.new("test/templates/temp_image.docx") do |r|

      # r.add_image("graphics1", File.join(Dir.pwd, 'test', 'templates', 'piriapolis.jpg'))
      # find_image_name_matches('Picture 1')
      r.add_image('IMAGE', File.join(Dir.pwd, 'test', 'templates', 'copy.jpeg'))

    end

    report.generate("test/result/test_word_image.docx")
