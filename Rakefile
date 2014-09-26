require "bundler/gem_tasks"
require 'launchy'
task :test do
  Dir.glob('./test/*_test.rb').each { |file| require file}
end

task :test_image do
  require './test/image_test.rb'
end

task :test_text do
  require './test/text_test.rb'
end

task :open do
  Dir.glob('./test/result/*.odt').each { |file| Launchy.open(file) }
end
