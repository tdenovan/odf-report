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

task :test_cchart do
  require './test/test_cchart.rb'
end

task :test_bchart do
  require './test/test_bchart.rb'
end

task :test_wchart do
  require './test/test_wchart.rb'
end

task :test_all do
  require './test/test_all.rb'
end

task :test_sfs do
  require './test/test_sfs.rb'
end

task :open do
  Dir.glob('./test/result/*.odt').each { |file| Launchy.open(file) }
end
