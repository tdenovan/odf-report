require "bundler/gem_tasks"
require 'launchy'
task :test do
  Dir.glob('./test/*_test.rb').each { |file| require file}
end

task :test_text do
  require './test/test_text.rb'
end

task :test_image do
  require './test/test_image.rb'
end

task :test_table do
  require './test/test_table.rb'
end

task :test_pchart do
  require './test/test_pchart.rb'
end

task :test_all do
  require './test/test_all.rb'
end

task :test_ugly_table do
  require './test/test_ugly_table.rb'
end

task :test_ugly_image do
  require './test/test_ugly_image.rb'
end

task :test_ugly_all do
  require './test/test_ugly_all.rb'
end

task :open do
  Dir.glob('./test/result/*.odt').each { |file| Launchy.open(file) }
end
