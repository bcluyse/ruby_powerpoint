# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'ruby_powerpoint/version'

Gem::Specification.new do |spec|
  spec.name          = "ruby_powerpoint"
  spec.version       = RubyPowerpoint::VERSION
  spec.authors       = ["bcluyse", "pythonicrubyist"]
  spec.email         = ["cluysebernard@gmail.com"]
  spec.description   = %q{A Ruby gem that can extract text and images from PowerPoint (pptx) files.}
  spec.summary       = %q{ruby_powerpoint is a Ruby gem that can extract title, content and images from Powerpont (pptx) slides.}
  spec.homepage      = "https://github.com/pythonicrubyist/ruby_powerpoint"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.required_ruby_version = '>= 1.9.2'

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
  spec.add_development_dependency 'rspec', '~> 2.13.0'

  spec.add_dependency 'nokogiri', '~> 1.6.0'
  spec.add_dependency 'rubyzip', '~> 0.9.9'  
end
