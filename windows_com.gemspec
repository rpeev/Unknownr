require 'rake'

require_relative 'lib/windows_com'

Gem::Specification.new do |spec|
  spec.name = 'windows_com'
  spec.version = WINDOWS_COM_VERSION

  spec.summary = 'Ruby FFI (x86) bindings to essential COM related Windows APIs'
  spec.description = 'Ruby FFI (x86) bindings to essential COM related Windows APIs'
  spec.homepage = 'https://github.com/rpeev/windows_com'

  spec.authors = ['Radoslav Peev']
  spec.email = ['rpeev@ymail.com']
  spec.licenses = ['MIT']

  spec.files = FileList[
    'LICENSE',
    'README.md',
    'RELNOTES.md',
    'lib/windows_com.rb',
    'lib/windows_com/*.rb',
    'examples/*.*'
  ]
  spec.require_paths = ['lib']
  spec.add_runtime_dependency('ffi', '~> 1')
end
