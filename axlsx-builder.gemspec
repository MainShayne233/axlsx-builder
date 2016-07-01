# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'axlsx/builder/version'

Gem::Specification.new do |spec|
  spec.name          = "axlsx-builder"
  spec.version       = Axlsx::Builder::VERSION
  spec.authors       = ["MainShayne233"]
  spec.email         = ["shaynetremblay@hotmail.com"]

  spec.summary       = 'An extension of Axlsx that allows you create spreadsheets blueprints for '\
                       'easy sheet generation, manipulation, and data input.'
  spec.homepage      = 'https://github.com/MainShayne233/axlsx-builder'
  spec.license       = 'MIT'


  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_development_dependency 'bundler', '~> 1.12'
  spec.add_development_dependency 'rake', '~> 10.0'
  spec.add_development_dependency 'rspec', '~> 3.0'
  spec.add_dependency 'axlsx', '~> 2.0.1.pre'
end
