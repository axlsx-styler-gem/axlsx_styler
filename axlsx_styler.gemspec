# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'axlsx_styler/version'

Gem::Specification.new do |spec|
  spec.name          = "axlsx_styler"
  spec.version       = AxlsxStyler::VERSION
  spec.authors       = ["Anton Sakovich"]
  spec.email         = ["sakovias@gmail.com"]
  spec.summary       = %q{Separate data from styles when using Axlsx gem.}
  # spec.description   = %q{TODO: Write a longer description. Optional.}
  spec.homepage      = ""
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_dependency "axlsx", "~> 2.0"
  spec.add_dependency "activesupport", "~> 3.1"

  spec.add_development_dependency "bundler", "~> 1.6"
  spec.add_development_dependency "rake", "~> 0.9"
  spec.add_development_dependency "minitest", "~> 5.0"
end
