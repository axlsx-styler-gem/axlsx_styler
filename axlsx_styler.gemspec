# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'axlsx_styler/version'

Gem::Specification.new do |spec|
  spec.name          = 'axlsx_styler'
  spec.version       = AxlsxStyler::VERSION
  spec.authors       = ['Anton Sakovich']
  spec.email         = ['sakovias@gmail.com']
  spec.summary       = %q(This gem allows to separate data from styles when using Axlsx gem.)
  spec.description   = %q(
    Axlsx gem is an excellent tool to build Excel worksheets. The sheets are
    created row-by-row and styles are immediately added to each cell when a
    row is created. This gem allows to follow an alternative route: fill out
    a spreadsheet with data and apply styles later. Styles can be added
    to individual cells as well as to ranges of cells. As a bonus, this gem
    also simplifies drawing borders around groups of cells.
  )
  spec.homepage      = 'https://github.com/sakovias/axlsx_styler'
  spec.license       = 'MIT'

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ['lib']
  spec.required_ruby_version = '>= 1.9.3'

  spec.add_dependency 'axlsx', '>= 2.0'
  spec.add_dependency 'activesupport', '>= 3.1'

  spec.add_development_dependency 'bundler', '~> 1.6'
  spec.add_development_dependency 'rake', '~> 0.9'
  spec.add_development_dependency 'minitest', '~> 5.0'
  spec.add_development_dependency 'awesome_print', '~> 1.6'
end
