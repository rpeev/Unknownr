require_relative 'lib/unknownr'
require 'rake/testtask'

Rake::TestTask.new do |t|
  t.test_files = FileList[
    'test/*_test.rb'
  ]
end

desc 'Build gem'
task :build => [:test] do |t|
  system "gem build unknownr.gemspec"
end

desc 'Push gem'
task :push => [:build] do |t|
  system "gem push unknownr-#{Unknownr::VERSION}.gem"
end

task :default => [:test]

if __FILE__ == $0
  Rake::Task[:default].invoke
end
