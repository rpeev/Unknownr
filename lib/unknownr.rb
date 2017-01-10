require 'ffi'

module Unknownr
	VERSION = '0.2.2'

	module Windows
		extend FFI::Library

		ffi_lib 'ole32'
		ffi_convention :stdcall

		S_OK = 0
		S_FALSE = 1

		E_UNEXPECTED = 0x8000FFFF - 0x1_0000_0000
		E_NOTIMPL = 0x80004001 - 0x1_0000_0000
		E_OUTOFMEMORY = 0x8007000E - 0x1_0000_0000
		E_INVALIDARG = 0x80070057 - 0x1_0000_0000
		E_NOINTERFACE = 0x80004002 - 0x1_0000_0000
		E_POINTER = 0x80004003 - 0x1_0000_0000
		E_HANDLE = 0x80070006 - 0x1_0000_0000
		E_ABORT = 0x80004004 - 0x1_0000_0000
		E_FAIL = 0x80004005 - 0x1_0000_0000
		E_ACCESSDENIED = 0x80070005 - 0x1_0000_0000
		E_PENDING = 0x8000000A - 0x1_0000_0000

		FACILITY_WIN32 = 7

		ERROR_CANCELLED = 1223

		def SUCCEEDED(hr) hr >= 0 end
		def FAILED(hr) hr < 0 end
		def HRESULT_FROM_WIN32(x) (x <= 0) ? x : (x & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000 end

		module_function \
			:SUCCEEDED,
			:FAILED,
			:HRESULT_FROM_WIN32

		def DetonateHresult(name, *args)
			failed = FAILED(result = send(name, *args)) and raise "#{name} failed (hresult #{format('%#08x', result)})."

			result
		ensure
			yield failed if block_given?
		end

		module_function :DetonateHresult

		class GUID < FFI::Struct
			layout \
				:Data1, :ulong,
				:Data2, :ushort,
				:Data3, :ushort,
				:Data4, [:uchar, 8]

			def self.[](s)
				raise 'Bad GUID format.' unless s =~ /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i

				new.tap { |guid|
					guid[:Data1] = s[0, 8].to_i(16)
					guid[:Data2] = s[9, 4].to_i(16)
					guid[:Data3] = s[14, 4].to_i(16)
					guid[:Data4][0] = s[19, 2].to_i(16)
					guid[:Data4][1] = s[21, 2].to_i(16)
					s[24, 12].split('').each_slice(2).with_index { |a, i|
						guid[:Data4][i + 2] = a.join('').to_i(16)
					}
				}
			end

			def ==(other) Windows.memcmp(other, self, size) == 0 end
		end

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

		module COM
			module Interface
				def self.[](*args)
					spec, iid, *ifaces = args.reverse

					spec.each { |name, signature| signature[0].unshift(:pointer) }

					Class.new(FFI::Struct) {
						const_set(:IID, iid)

						const_set(:VTBL, Class.new(FFI::Struct) {
							const_set(:SPEC, Hash[(ifaces.map { |iface| iface::VTBL::SPEC.to_a } << spec.to_a).flatten(1)])

							layout \
								*self::SPEC.map { |name, signature| [name, callback(*signature)] }.flatten
						})

						layout \
							:lpVtbl, :pointer
					}
				end
			end

			module Helpers
				def QueryInstance(klass)
					instance = nil

					FFI::MemoryPointer.new(:pointer) { |ppv|
						QueryInterface(klass::IID, ppv)

						instance = klass.new(ppv.read_pointer)
					}

					begin
						yield instance; return self
					ensure
						instance.Release
					end if block_given?

					instance
				end

				def UseInstance(klass, name, *args)
					instance = nil

					FFI::MemoryPointer.new(:pointer) { |ppv|
						send(name, *args, klass::IID, ppv)

						yield instance = klass.new(ppv.read_pointer)
					}

					self
				ensure
					instance.Release if instance
				end
			end

			module Instance
				def self.[](iface)
					Class.new(iface) {
						send(:include, Helpers)

						def initialize(pointer)
							self.pointer = pointer

							@vtbl = self.class::VTBL.new(self[:lpVtbl])
						end

						attr_reader :vtbl

						self::VTBL.members.each { |name|
							define_method(name) { |*args|
								raise "#{self}.#{name} failed." if Windows.FAILED(result = @vtbl[name].call(self, *args)); result
							}
						}
					}
				end
			end

			module Factory
				def self.[](iface, clsid)
					Class.new(iface) {
						send(:include, Helpers)

						const_set(:CLSID, clsid)

						def initialize(opts = {})
							@opts = opts

							@opts[:clsctx] ||= CLSCTX_INPROC_SERVER

							FFI::MemoryPointer.new(:pointer) { |ppv|
								raise "CoCreateInstance failed (#{self.class})." if
									Windows.FAILED(Windows.CoCreateInstance(self.class::CLSID, nil, @opts[:clsctx], self.class::IID, ppv))

								self.pointer = ppv.read_pointer
							}

							@vtbl = self.class::VTBL.new(self[:lpVtbl])
						end

						attr_reader :vtbl

						self::VTBL.members.each { |name|
							define_method(name) { |*args|
								raise "#{self}.#{name} failed." if Windows.FAILED(result = @vtbl[name].call(self, *args)); result
							}
						}
					}
				end
			end

			module Callback
				def self.[](iface)
					Class.new(iface) {
						send(:include, Helpers)

						def initialize(opts = {})
							@vtbl, @refc = self.class::VTBL.new, 1

							@vtbl.members.each { |name|
								@vtbl[name] = instance_variable_set("@fn#{name}",
									FFI::Function.new(*@vtbl.class::SPEC[name].reverse, convention: :stdcall) { |*args|
										send(name, *args[1..-1])
									}
								)
							}

							self[:lpVtbl] = @vtbl

							begin
								yield self
							ensure
								Release()
							end if block_given?
						end

						attr_reader :vtbl, :refc

						def QueryInterface(riid, ppv)
							if [IUnknown::IID, self.class::IID].any? { |iid| Windows.memcmp(riid, iid, iid.size) == 0 }
								ppv.write_pointer(self)
							else
								ppv.write_pointer(0); return E_NOINTERFACE
							end

							AddRef(); S_OK
						end

						def AddRef
							@refc += 1
						end

						def Release
							@refc -= 1
						end

						(self::VTBL.members - IUnknown::VTBL.members).each { |name|
							define_method(name) { |*args|
								E_NOTIMPL
							}
						}
					}
				end
			end
		end

		IUnknown = COM::Interface[
			GUID['00000000-0000-0000-C000-000000000046'],

			QueryInterface: [[:pointer, :pointer], :long],
			AddRef: [[], :ulong],
			Release: [[], :ulong]
		]

		Unknown = COM::Instance[IUnknown]

		IDispatch = COM::Interface[IUnknown,
			GUID['00020400-0000-0000-C000-000000000046'],

			GetTypeInfoCount: [[:pointer], :long],
			GetTypeInfo: [[:uint, :ulong, :pointer], :long],
			GetIDsOfNames: [[:pointer, :pointer, :uint, :ulong, :pointer], :long],
			Invoke: [[:long, :pointer, :ulong, :ushort, :pointer, :pointer, :pointer, :pointer], :long]
		]

		Dispatch = COM::Instance[IDispatch]
		DCallback = COM::Callback[IDispatch]

		IConnectionPointContainer = COM::Interface[IUnknown,
			GUID['B196B284-BAB4-101A-B69C-00AA00341D07'],

			EnumConnectionPoints: [[:pointer], :long],
			FindConnectionPoint: [[:pointer, :pointer], :long]
		]

		ConnectionPointContainer = COM::Instance[IConnectionPointContainer]

		IConnectionPoint = COM::Interface[IUnknown,
			GUID['B196B286-BAB4-101A-B69C-00AA00341D07'],

			GetConnectionInterface: [[:pointer], :long],
			GetConnectionPointContainer: [[:pointer], :long],
			Advise: [[:pointer, :pointer], :long],
			Unadvise: [[:ulong], :long],
			EnumConnections: [[:pointer], :long]
		]

		ConnectionPoint = COM::Instance[IConnectionPoint]

		IObjectWithSite = COM::Interface[IUnknown,
			GUID['FC4801A3-2BA9-11CF-A229-00AA003D7352'],

			SetSite: [[:pointer], :long],
			GetSite: [[:pointer, :pointer], :long]
		]

		ObjectWithSite = COM::Callback[IObjectWithSite]

		attach_function :OleInitialize, [:pointer], :long
		attach_function :OleUninitialize, [], :void

		def InitializeOle
			DetonateHresult(:OleInitialize, nil)

			at_exit { OleUninitialize() }
		end

		module_function :InitializeOle

		attach_function :CoTaskMemAlloc, [:ulong], :pointer
		attach_function :CoTaskMemFree, [:pointer], :void

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

		class PROPERTYKEY < FFI::Struct
			layout \
				:fmtid, GUID,
				:pid, :ulong

			def self.[](type, index)
				new.tap { |key|
					key[:fmtid].tap { |guid|
						guid[:Data1] = 0x00000000 + index
						guid[:Data2] = 0x7363
						guid[:Data3] = 0x696e
						[0x84, 0x41, 0x79, 0x8a, 0xcf, 0x5a, 0xeb, 0xb7].each_with_index { |part, i|
							guid[:Data4][i] = part
						}
					}

					key[:pid] = type
				}
			end

			def ==(other) Windows.memcmp(other, self, size) == 0 end
		end

		module AnonymousSupport
			def [](k)
				if members.include?(k)
					super
				elsif self[:_].members.include?(k)
					self[:_][k]
				else
					self[:_][:_][k]
				end
			end

			def []=(k, v)
				if members.include?(k)
					super
				elsif self[:_].members.include?(k)
					self[:_][k] = v
				else
					self[:_][:_][k] = v
				end
			end
		end

		class LARGE_INTEGER < FFI::Union
			include AnonymousSupport

			layout \
				:_, Class.new(FFI::Struct) {
					layout \
						:LowPart, :ulong,
						:HighPart, :long
				},

				:QuadPart, :long_long
		end

		class ULARGE_INTEGER < FFI::Union
			include AnonymousSupport

			layout \
				:_, Class.new(FFI::Struct) {
					layout \
						:LowPart, :ulong,
						:HighPart, :ulong
				},

				:QuadPart, :ulong_long
		end

		class FILETIME < FFI::Struct
			layout \
				:dwLowDateTime, :ulong,
				:dwHighDateTime, :ulong
		end

		class BSTRBLOB < FFI::Struct
			layout \
				:cbSize, :ulong,
				:pData, :pointer
		end

		class BLOB < FFI::Struct
			layout \
				:cbSize, :ulong,
				:pBlobData, :pointer
		end

		class CA < FFI::Struct
			layout \
				:cElems, :ulong,
				:pElems, :pointer
		end

		class DECIMAL < FFI::Struct
			layout \
				:wReserved, :ushort,
				:scale, :uchar,
				:sign, :uchar,
				:Hi32, :ulong,
				:Lo64, :ulong_long
		end

		class VARIANT < FFI::Union
			include AnonymousSupport

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

		class PROPVARIANT < FFI::Union
			include AnonymousSupport

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

			def self.[](t, *v) new.tap { |var| var.send("#{t}=", *v) } end

			def bool; raise 'Wrong type tag.' unless self[:vt] == VT_BOOL; self[:boolVal] != 0 end
			def bool=(bool) self[:vt] = VT_BOOL; self[:boolVal] = (bool) ? -1 : 0 end

			def int; raise 'Wrong type tag.' unless self[:vt] == VT_I4; self[:intVal] end
			def int=(int) self[:vt] = VT_I4; self[:intVal] = int end

			def uint; raise 'Wrong type tag.' unless self[:vt] == VT_UI4; self[:uintVal] end
			def uint=(uint) self[:vt] = VT_UI4; self[:uintVal] = uint end

			def unknown
				raise 'Wrong type tag.' unless self[:vt] == VT_UNKNOWN

				yield Unknown.new(self[:punkVal])
			ensure
				Windows.PropVariantClear(self)
			end

			def unknown=(unknown) self[:vt] = VT_UNKNOWN; self[:punkVal] = unknown.pointer; unknown.AddRef end

			def wstring; raise 'Wrong type tag.' unless self[:vt] == VT_LPWSTR; Windows.WCSTOMBS(self[:pwszVal]) end

			def wstring=(string)
				self[:vt] = VT_LPWSTR

				FFI::MemoryPointer.new(:pointer) { |p|
					Windows.DetonateHresult(:SHStrDup, string, p)

					self[:pwszVal] = p.read_pointer
				}
			end

			def decimal
				raise 'Wrong type tag.' unless self[:vt] == VT_DECIMAL

				Rational(self[:decVal][:Lo64], 10 ** self[:decVal][:scale]) + self[:decVal][:Hi32]
			end

			def decimal=(decimal) self[:vt] = VT_DECIMAL; self[:decVal][:Lo64] = decimal end
		end

		attach_function :PropVariantClear, [:pointer], :long

		class SAFEARRAYBOUND < FFI::Struct
			layout \
				:cElements, :ulong,
				:lLbound, :long
		end

		class SAFEARRAY < FFI::Struct
			layout \
				:cDims, :ushort,
				:fFeatures, :ushort,
				:cbElements, :ulong,
				:cLocks, :ulong,
				:pvData, :pointer,
				:rgsabound, [SAFEARRAYBOUND, 1]
		end

		ffi_lib 'oleaut32'
		ffi_convention :stdcall

		attach_function :SysAllocString, [:buffer_in], :pointer
		attach_function :SysFreeString, [:pointer], :void
		attach_function :SysStringLen, [:pointer], :uint

		attach_function :SafeArrayCreateVector, [:ushort, :long, :uint], :pointer
		attach_function :SafeArrayDestroy, [:pointer], :long
		attach_function :SafeArrayAccessData, [:pointer, :pointer], :long
		attach_function :SafeArrayUnaccessData, [:pointer], :long

		OLEIVERB_PRIMARY = 0
		OLEIVERB_SHOW = -1
		OLEIVERB_OPEN = -2
		OLEIVERB_HIDE = -3
		OLEIVERB_UIACTIVATE = -4
		OLEIVERB_INPLACEACTIVATE = -5
		OLEIVERB_DISCARDUNDOSTATE = -6

		IOleWindow = COM::Interface[IUnknown,
			GUID['00000114-0000-0000-C000-000000000046'],

			GetWindow: [[:pointer], :long],
			ContextSensitiveHelp: [[:int], :long]
		]

		OleWindow = COM::Instance[IOleWindow]

		IOleInPlaceObject = COM::Interface[IOleWindow,
			GUID['00000113-0000-0000-C000-000000000046'],

			InPlaceDeactivate: [[], :long],
			UIDeactivate: [[], :long],
			SetObjectRects: [[:pointer, :pointer], :long],
			ReactivateAndUndo: [[], :long]
		]

		OleInPlaceObject = COM::Instance[IOleInPlaceObject]

		IOleInPlaceSite = COM::Interface[IOleWindow,
			GUID['00000119-0000-0000-C000-000000000046'],

			CanInPlaceActivate: [[], :long],
			OnInPlaceActivate: [[], :long],
			OnUIActivate: [[], :long],
			GetWindowContext: [[:pointer, :pointer, :pointer, :pointer, :pointer], :long],
			Scroll: [[:long_long], :long],
			OnUIDeactivate: [[:int], :long],
			OnInPlaceDeactivate: [[], :long],
			DiscardUndoState: [[], :long],
			DeactivateAndUndo: [[], :long],
			OnPosRectChange: [[:pointer], :long]
		]

		OleInPlaceSite = COM::Callback[IOleInPlaceSite]

		IOleClientSite = COM::Interface[IUnknown,
			GUID['00000118-0000-0000-C000-000000000046'],

			SaveObject: [[], :long],
			GetMoniker: [[:ulong, :ulong, :pointer], :long],
			GetContainer: [[:pointer], :long],
			ShowObject: [[], :long],
			OnShowWindow: [[:int], :long],
			RequestNewObjectLayout: [[], :long]
		]

		OleClientSite = COM::Callback[IOleClientSite]

		OLEGETMONIKER_ONLYIFTHERE = 1
		OLEGETMONIKER_FORCEASSIGN = 2
		OLEGETMONIKER_UNASSIGN = 3
		OLEGETMONIKER_TEMPFORUSER = 4

		OLEWHICHMK_CONTAINER = 1
		OLEWHICHMK_OBJREL = 2
		OLEWHICHMK_OBJFULL = 3

		USERCLASSTYPE_FULL = 1
		USERCLASSTYPE_SHORT = 2
		USERCLASSTYPE_APPNAME = 3

		OLEMISC_RECOMPOSEONRESIZE = 0x00000001
		OLEMISC_ONLYICONIC = 0x00000002
		OLEMISC_INSERTNOTREPLACE = 0x00000004
		OLEMISC_STATIC = 0x00000008
		OLEMISC_CANTLINKINSIDE = 0x00000010
		OLEMISC_CANLINKBYOLE1 = 0x00000020
		OLEMISC_ISLINKOBJECT = 0x00000040
		OLEMISC_INSIDEOUT = 0x00000080
		OLEMISC_ACTIVATEWHENVISIBLE = 0x00000100
		OLEMISC_RENDERINGISDEVICEINDEPENDENT = 0x00000200
		OLEMISC_INVISIBLEATRUNTIME = 0x00000400
		OLEMISC_ALWAYSRUN = 0x00000800
		OLEMISC_ACTSLIKEBUTTON = 0x00001000
		OLEMISC_ACTSLIKELABEL = 0x00002000
		OLEMISC_NOUIACTIVATE = 0x00004000
		OLEMISC_ALIGNABLE = 0x00008000
		OLEMISC_SIMPLEFRAME = 0x00010000
		OLEMISC_SETCLIENTSITEFIRST = 0x00020000
		OLEMISC_IMEMODE = 0x00040000
		OLEMISC_IGNOREACTIVATEWHENVISIBLE = 0x00080000
		OLEMISC_WANTSTOMENUMERGE = 0x00100000
		OLEMISC_SUPPORTSMULTILEVELUNDO = 0x00200000

		OLECLOSE_SAVEIFDIRTY = 0
		OLECLOSE_NOSAVE = 1
		OLECLOSE_PROMPTSAVE = 2

		IOleObject = COM::Interface[IUnknown,
			GUID['00000112-0000-0000-C000-000000000046'],

			SetClientSite: [[:pointer], :long],
			GetClientSite: [[:pointer], :long],
			SetHostNames: [[:pointer, :pointer], :long],
			Close: [[:ulong], :long],
			SetMoniker: [[:ulong, :pointer], :long],
			GetMoniker: [[:ulong, :ulong, :pointer], :long],
			InitFromData: [[:pointer, :int, :ulong], :long],
			GetClipboardData: [[:ulong, :pointer], :long],
			DoVerb: [[:long, :pointer, :pointer, :long, :pointer, :pointer], :long],
			EnumVerbs: [[:pointer], :long],
			Update: [[], :long],
			IsUpToDate: [[], :long],
			GetUserClassID: [[:pointer], :long],
			GetUserType: [[:ulong, :pointer], :long],
			SetExtent: [[:ulong, :pointer], :long],
			GetExtent: [[:ulong, :pointer], :long],
			Advise: [[:pointer, :pointer], :long],
			Unadvise: [[:ulong], :long],
			EnumAdvise: [[:pointer], :long],
			GetMiscStatus: [[:ulong, :pointer], :long],
			SetColorScheme: [[:pointer], :long]
		]

		OleObject = COM::Instance[IOleObject]

		class PARAMDATA < FFI::Struct
			layout \
				:szName, :pointer,
				:vt, :ushort
		end

		CC_FASTCALL = 0
		CC_CDECL = 1
		CC_MSCPASCAL = CC_CDECL + 1
		CC_PASCAL = CC_MSCPASCAL
		CC_MACPASCAL = CC_PASCAL + 1
		CC_STDCALL = CC_MACPASCAL + 1
		CC_FPFASTCALL = CC_STDCALL + 1
		CC_SYSCALL = CC_FPFASTCALL + 1
		CC_MPWCDECL = CC_SYSCALL + 1
		CC_MPWPASCAL = CC_MPWCDECL + 1
		CC_MAX = CC_MPWPASCAL + 1

		DISPATCH_METHOD = 0x1
		DISPATCH_PROPERTYGET = 0x2
		DISPATCH_PROPERTYPUT = 0x4
		DISPATCH_PROPERTYPUTREF = 0x8

		class METHODDATA < FFI::Struct
			layout \
				:szName, :pointer,
				:ppdata, :pointer,
				:dispid, :long,
				:iMeth, :uint,
				:cc, :uint,
				:cArgs, :uint,
				:wFlags, :ushort,
				:vtReturn, :ushort
		end

		class INTERFACEDATA < FFI::Struct
			layout \
				:pmethdata, :pointer,
				:cMembers, :uint
		end

		attach_function :CreateDispTypeInfo, [:pointer, :ulong, :pointer], :long

		class DISPPARAMS < FFI::Struct
			layout \
				:rgvarg, :pointer,
				:rgdispidNamedArgs, :pointer,
				:cArgs, :uint,
				:cNamedArgs, :uint
		end

		attach_function :DispInvoke, [:pointer, :pointer, :long, :ushort, :pointer, :pointer, :pointer, :pointer], :long
	end
end

if __FILE__ == $0
	puts Unknownr::VERSION
end
