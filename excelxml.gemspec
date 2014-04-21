# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'excelxml/version'

Gem::Specification.new do |spec|
  spec.name          = "excelxml"
  spec.version       = ExcelXml::VERSION
  spec.authors       = ["Ben Delsol"]
  spec.summary       = %q{Parses the data out of Excel XML 2003 workbooks/sheets.}
  spec.description   = %q{}
  spec.homepage      = "https://github.com/bdiz/excelxml"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_dependency "happymapper"

  spec.add_development_dependency "bundler", "~> 1.5"
  spec.add_development_dependency "rake"
end
