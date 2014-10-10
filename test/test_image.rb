require './lib/odf-report'
require 'faker'

report = ODFReport::Report.new("test/templates/temp_image.docx") do |r|

  r.add_image('image_01', File.join(Dir.pwd, 'test', 'templates', 'replace_01.jpg')) # Image with the same title
  r.add_image('image_01', File.join(Dir.pwd, 'test', 'templates', 'replace_02.jpg')) # Image with the same title
  r.add_image('image_02', File.join(Dir.pwd, 'test', 'templates', 'replace_03.jpg')) # Check for upper/lowercase
  r.add_image('image_03', File.join(Dir.pwd, 'test', 'templates', 'replace_04.jpg')) # Image source in word is the same
  r.add_image('image_04', File.join(Dir.pwd, 'test', 'templates', 'replace_05.jpg')) # Image source in word is the same
  r.add_image('image_05', File.join(Dir.pwd, 'test', 'templates', 'replace_06.jpg')) # Image_05 does NOT exist

end

report.generate("test/result/test_word_image.docx")
