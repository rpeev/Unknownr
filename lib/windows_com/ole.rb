if __FILE__ == $0
	require_relative 'common'
	require_relative 'libc'
	require_relative 'shlwapi'
end

module WindowsCOM
	ffi_lib 'ole32'
	ffi_convention :stdcall

	attach_function :OleInitialize, [:pointer], :long
	attach_function :OleUninitialize, [], :void

	def InitializeOle
		DetonateHresult(:OleInitialize, nil)

		STDERR.puts "OLE initialized" if $DEBUG

		at_exit {
			OleUninitialize()

			STDERR.puts "OLE uninitialized" if $DEBUG
		}
	end

	module_function \
		:InitializeOle

	InitializeOle() if WINDOWS_COM_OLE_INIT

	attach_function :CoTaskMemAlloc, [:ulong], :pointer
	attach_function :CoTaskMemFree, [:pointer], :void

	class LARGE_INTEGER < FFI::Union
		include FFIStructAnonymousAccess

		layout \
			:_, Class.new(FFI::Struct) {
				layout \
					:LowPart, :ulong,
					:HighPart, :long
			},

			:QuadPart, :long_long
	end

	class ULARGE_INTEGER < FFI::Union
		include FFIStructAnonymousAccess

		layout \
			:_, Class.new(FFI::Struct) {
				layout \
					:LowPart, :ulong,
					:HighPart, :ulong
			},

			:QuadPart, :ulong_long
	end

	class DECIMAL < FFI::Struct
		layout \
			:wReserved, :ushort,
			:scale, :uchar,
			:sign, :uchar,
			:Hi32, :ulong,
			:Lo64, :ulong_long
	end

	class BLOB < FFI::Struct
		layout \
			:cbSize, :ulong,
			:pBlobData, :pointer
	end

	class BSTRBLOB < FFI::Struct
		layout \
			:cbSize, :ulong,
			:pData, :pointer
	end

	class FILETIME < FFI::Struct
		layout \
			:dwLowDateTime, :ulong,
			:dwHighDateTime, :ulong
	end

	class CA < FFI::Struct
		layout \
			:cElems, :ulong,
			:pElems, :pointer
	end

	VT_EMPTY = 0
	VT_NULL = 1
	VT_I2 = 2
	VT_I4 = 3
	VT_R4 = 4
	VT_R8 = 5
	VT_CY = 6
	VT_DATE = 7
	VT_BSTR = 8
	VT_DISPATCH = 9
	VT_ERROR = 10
	VT_BOOL = 11
	VT_VARIANT = 12
	VT_UNKNOWN = 13
	VT_DECIMAL = 14
	VT_I1 = 16
	VT_UI1 = 17
	VT_UI2 = 18
	VT_UI4 = 19
	VT_I8 = 20
	VT_UI8 = 21
	VT_INT = 22
	VT_UINT = 23
	VT_VOID = 24
	VT_HRESULT = 25
	VT_PTR = 26
	VT_SAFEARRAY = 27
	VT_CARRAY = 28
	VT_USERDEFINED = 29
	VT_LPSTR = 30
	VT_LPWSTR = 31
	VT_FILETIME = 64
	VT_BLOB = 65
	VT_STREAM = 66
	VT_STORAGE = 67
	VT_STREAMED_OBJECT = 68
	VT_STORED_OBJECT = 69
	VT_BLOB_OBJECT = 70
	VT_CF = 71
	VT_CLSID = 72
	VT_VECTOR = 0x1000
	VT_ARRAY = 0x2000
	VT_BYREF = 0x4000
	VT_RESERVED = 0x8000
	VT_ILLEGAL = 0xffff
	VT_ILLEGALMASKED = 0xfff
	VT_TYPEMASK = 0xff

	class VARIANT < FFI::Union
		extend VariantBasicCreation
		include FFIStructAnonymousAccess

		layout \
			:_, Class.new(FFI::Struct) {
				layout \
					:vt, :ushort,
					:wReserved1, :ushort,
					:wReserved2, :ushort,
					:wReserved3, :ushort,
					:_, Class.new(FFI::Union) {
						layout \
							:llVal, :long_long,
							:lVal, :long,
							:bVal, :uchar,
							:iVal, :short,
							:fltVal, :float,
							:dblVal, :double,
							:boolVal, :short,
							:bool, :short,
							:scode, :long,
							:cyVal, :ulong_long,
							:date, :double,
							:bstrVal, :pointer,
							:punkVal, :pointer,
							:pdispVal, :pointer,
							:parray, :pointer,
							:pbVal, :pointer,
							:piVal, :pointer,
							:plVal, :pointer,
							:pllVal, :pointer,
							:pfltVal, :pointer,
							:pdblVal, :pointer,
							:pboolVal, :pointer,
							:pbool, :pointer,
							:pscode, :pointer,
							:pcyVal, :pointer,
							:pdate, :pointer,
							:pbstrVal, :pointer,
							:ppunkVal, :pointer,
							:ppdispVal, :pointer,
							:pparrayv, :pointer,
							:pvarVal, :pointer,
							:byref, :pointer,
							:cVal, :char,
							:uiVal, :ushort,
							:ulVal, :ulong,
							:ullVal, :ulong_long,
							:intVal, :int,
							:uintVal, :uint,
							:pdecVal, :pointer,
							:pcVal, :pointer,
							:puiVal, :pointer,
							:pulVal, :pointer,
							:pullVal, :pointer,
							:pintVal, :pointer,
							:puintVal, :pointer,
							:BRECORD, Class.new(FFI::Struct) {
								layout \
									:pvRecord, :pointer,
									:pRecInfo, :pointer
							}
					}
			},

			:decVal, DECIMAL
	end

	class PROPERTYKEY < FFI::Struct
		include FFIStructMemoryEquality

		layout \
			:fmtid, GUID,
			:pid, :ulong

		def self.[](type, index)
			propkey = new

			propkey[:pid] = type

			guid = propkey[:fmtid]
			guid[:Data1] = 0x00000000 + index
			guid[:Data2] = 0x7363
			guid[:Data3] = 0x696e
			[0x84, 0x41, 0x79, 0x8a, 0xcf, 0x5a, 0xeb, 0xb7].each_with_index { |part, i|
				guid[:Data4][i] = part
			}

			propkey
		end
	end

	class PROPVARIANT < FFI::Union
		extend VariantBasicCreation
		include FFIStructAnonymousAccess

		layout \
			:_, Class.new(FFI::Struct) {
				layout \
					:vt, :ushort,
					:wReserved1, :ushort,
					:wReserved2, :ushort,
					:wReserved3, :ushort,
					:_, Class.new(FFI::Union) {
						layout \
							:cVal, :char,
							:bVal, :uchar,
							:iVal, :short,
							:uiVal, :ushort,
							:lVal, :long,
							:ulVal, :ulong,
							:intVal, :int,
							:uintVal, :uint,
							:hVal, LARGE_INTEGER,
							:uhVal, ULARGE_INTEGER,
							:fltVal, :float,
							:dblVal, :double,
							:boolVal, :short,
							:bool, :short,
							:scode, :long,
							:cyVal, :long_long,
							:date, :double,
							:filetime, FILETIME,
							:puuid, :pointer,
							:pclipdata, :pointer,
							:bstrVal, :pointer,
							:bstrblobVal, BSTRBLOB,
							:blob, BLOB,
							:pszVal, :pointer,
							:pwszVal, :pointer,
							:punkVal, :pointer,
							:pdispVal, :pointer,
							:pStream, :pointer,
							:pStorage, :pointer,
							:pVersionedStream, :pointer,
							:parray, :pointer,
							:cac, CA,
							:caub, CA,
							:cai, CA,
							:caui, CA,
							:cal, CA,
							:caul, CA,
							:cah, CA,
							:cauh, CA,
							:caflt, CA,
							:cadbl, CA,
							:cabool, CA,
							:cascode, CA,
							:cacy, CA,
							:cadate, CA,
							:cafiletime, CA,
							:cauuid, CA,
							:caclipdata, CA,
							:cabstr, CA,
							:cabstrblob, CA,
							:calpstr, CA,
							:calpwstr, CA,
							:capropvar, CA,
							:pcVal, :pointer,
							:pbVal, :pointer,
							:piVal, :pointer,
							:puiVal, :pointer,
							:plVal, :pointer,
							:pulVal, :pointer,
							:pintVal, :pointer,
							:puintVal, :pointer,
							:pfltVal, :pointer,
							:pdblVal, :pointer,
							:pboolVal, :pointer,
							:pdecVal, :pointer,
							:pscode, :pointer,
							:pcyVal, :pointer,
							:pdate, :pointer,
							:pbstrVal, :pointer,
							:ppunkVal, :pointer,
							:ppdispVal, :pointer,
							:pparray, :pointer,
							:pvarVal, :pointer
					}
			},

			:decVal, DECIMAL
	end

	attach_function :PropVariantClear, [:pointer], :long

	CLSCTX_INPROC_SERVER = 0x1
	CLSCTX_INPROC_HANDLER = 0x2
	CLSCTX_LOCAL_SERVER = 0x4
	CLSCTX_INPROC_SERVER16 = 0x8
	CLSCTX_REMOTE_SERVER = 0x10
	CLSCTX_INPROC_HANDLER16 = 0x20
	CLSCTX_RESERVED1 = 0x40
	CLSCTX_RESERVED2 = 0x80
	CLSCTX_RESERVED3 = 0x100
	CLSCTX_RESERVED4 = 0x200
	CLSCTX_NO_CODE_DOWNLOAD = 0x400
	CLSCTX_RESERVED5 = 0x800
	CLSCTX_NO_CUSTOM_MARSHAL = 0x1000
	CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000
	CLSCTX_NO_FAILURE_LOG = 0x4000
	CLSCTX_DISABLE_AAA = 0x8000
	CLSCTX_ENABLE_AAA = 0x10000
	CLSCTX_FROM_DEFAULT_CONTEXT = 0x20000
	CLSCTX_ACTIVATE_32_BIT_SERVER = 0x40000
	CLSCTX_ACTIVATE_64_BIT_SERVER = 0x80000
	CLSCTX_ENABLE_CLOAKING = 0x100000
	CLSCTX_PS_DLL = -0x80000000
	CLSCTX_INPROC = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER
	CLSCTX_ALL = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER
	CLSCTX_SERVER = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER

	attach_function :CoCreateInstance, [:pointer, :pointer, :ulong, :pointer, :pointer], :long

	IUnknown = COMInterface[nil,
		'00000000-0000-0000-C000-000000000046',

		QueryInterface: [[:pointer, :pointer], :long],
		AddRef: [[], :ulong],
		Release: [[], :ulong]
	]

	IDispatch = COMInterface[IUnknown,
		'00020400-0000-0000-C000-000000000046',

		GetTypeInfoCount: [[:pointer], :long],
		GetTypeInfo: [[:uint, :ulong, :pointer], :long],
		GetIDsOfNames: [[:pointer, :pointer, :uint, :ulong, :pointer], :long],
		Invoke: [[:long, :pointer, :ulong, :ushort, :pointer, :pointer, :pointer, :pointer], :long]
	]

	IConnectionPointContainer = COMInterface[IUnknown,
		'B196B284-BAB4-101A-B69C-00AA00341D07',

		EnumConnectionPoints: [[:pointer], :long],
		FindConnectionPoint: [[:pointer, :pointer], :long]
	]

	IConnectionPoint = COMInterface[IUnknown,
		'B196B286-BAB4-101A-B69C-00AA00341D07',

		GetConnectionInterface: [[:pointer], :long],
		GetConnectionPointContainer: [[:pointer], :long],
		Advise: [[:pointer, :pointer], :long],
		Unadvise: [[:ulong], :long],
		EnumConnections: [[:pointer], :long]
	]

	IObjectWithSite = COMInterface[IUnknown,
		'FC4801A3-2BA9-11CF-A229-00AA003D7352',

		SetSite: [[:pointer], :long],
		GetSite: [[:pointer, :pointer], :long]
	]
end
