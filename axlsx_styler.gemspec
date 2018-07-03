lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)

require 'axlsx_styler/version'

Gem::Specification.new do |spec|
  spec.name          = 'axlsx_styler'
  spec.version       = AxlsxStyler::VERSION
  spec.authors       = ['Weston Ganger', 'Anton Sakovich']
  spec.email         = ['weston@westonganger.com', 'sakovias@gmail.com']
  spec.summary       = 'Build clean and maintainable styles for your axlsx spreadsheets. Build your spreadsheeet with data and then apply styles later.'
  spec.description   = spec.summary
  spec.homepage      = 'https://github.com/axlsx-styler-gem/axlsx_styler'
  spec.license       = 'MIT'

  spec.files         = `git ls-files -z`.split("\x0")
  spec.files         = Dir.glob("{lib/**/*}") + %w{ LICENSE.txt README.md Rakefile CHANGELOG.md }
  spec.test_files    = Dir.glob("{test/**/*}")
  spec.require_paths = ['lib']

  spec.required_ruby_version = '>= 1.9.3'

  spec.add_dependency 'axlsx', ['>= 2.0', '< 4']
  spec.add_dependency 'activesupport', '>= 3.1'

  spec.add_development_dependency 'bundler'
  spec.add_development_dependency 'rake'
  spec.add_development_dependency 'minitest'
  spec.add_development_dependency 'appraisal'
end
