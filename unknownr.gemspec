require File.expand_path('../lib/unknownr', __FILE__) # require_relative doesn't work here in ruby 1.9
require 'rake'

Gem::Specification.new do |s|
  s.name = 'unknownr'
  s.version = Unknownr::VERSION

  s.summary = 'Ruby-FFI (x86) bindings to essential COM-related Windows APIs'
  s.description = 'Ruby-FFI (x86) bindings to essential COM-related Windows APIs'
  s.homepage = 'https://github.com/rpeev/unknownr'

  s.authors = ['Radoslav Peev']
  s.email = ['rpeev@ymail.com']
  s.licenses = ['MIT']

  s.files = FileList[
    'LICENSE',
    'README.md',
    'RELNOTES.md',
    'lib/unknownr.rb',
    'examples/*.*'
  ]
  s.require_paths = ['lib']
  s.add_runtime_dependency('ffi', '~> 1')
end
