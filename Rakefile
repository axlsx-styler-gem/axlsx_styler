require "bundler/gem_tasks"

task :test do 
  require 'rake/testtask'

  Rake::TestTask.new(:test) do |t|
    t.libs << 'lib'
    t.libs << 'test'
    t.pattern = 'test/**/*_test.rb'
    t.verbose = false
  end
end

task default: :test

task :console do
  require 'axlsx_styler'

  require 'irb'
  binding.irb
end
