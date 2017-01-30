if __FILE__ == $0
	require_relative 'common'
end

module WindowsCOM
	ffi_lib FFI::Library::LIBC
	ffi_convention :cdecl

	attach_function :windows_com_memcmp, :memcmp, [
		:pointer,
		:pointer,
		:size_t
	], :int
end
