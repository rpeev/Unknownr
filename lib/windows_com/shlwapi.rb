if __FILE__ == $0
	require_relative 'common'
	require_relative 'libc'
end

module WindowsCOM
	ffi_lib 'shlwapi'
	ffi_convention :stdcall

	attach_function :SHStrDup, :SHStrDupW, [:string, :buffer_out], :long
end
