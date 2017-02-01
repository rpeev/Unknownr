require 'ffi'

WINDOWS_COM_VERSION = '1.0.0'

WINDOWS_COM_OLE_INIT = true unless defined?(WINDOWS_COM_OLE_INIT)

module WindowsCOM
	extend FFI::Library

	def DetonateHresult(name, *args)
		hresult = send(name, *args)
		failed = FAILED(hresult)

		raise "#{name} failed (hresult: #{format('%#08x', hresult)})" if failed

		hresult
	ensure
		yield hresult if failed && block_given?
	end

	module_function \
		:DetonateHresult

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

	def SUCCEEDED(hr)
		hr >= 0
	end

	def FAILED(hr)
		hr < 0
	end

	def HRESULT_FROM_WIN32(x)
		(x <= 0) ?
			x :
			(x & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000
	end

	module_function \
		:SUCCEEDED,
		:FAILED,
		:HRESULT_FROM_WIN32

	class GUID < FFI::Struct
		layout \
			:Data1, :ulong,
			:Data2, :ushort,
			:Data3, :ushort,
			:Data4, [:uchar, 8]

		def self.[](str)
			raise 'Bad GUID format' unless str =~ /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i

			guid = new

			guid[:Data1] = str[0, 8].to_i(16)
			guid[:Data2] = str[9, 4].to_i(16)
			guid[:Data3] = str[14, 4].to_i(16)
			guid[:Data4][0] = str[19, 2].to_i(16)
			guid[:Data4][1] = str[21, 2].to_i(16)
			str[24, 12].split('').each_slice(2).with_index { |a, i|
				guid[:Data4][i + 2] = a.join('').to_i(16)
			}

			guid
		end

		def ==(other)
			windows_com_memcmp(self, other, self.size) == 0
		end
	end

	class COMVptr_ < FFI::Struct
		layout \
			:lpVtbl, :pointer
	end

	module COMVtbl_
		def self.[](parent_vtbl, spec)
			spec.each { |name, sig|
				sig[0].unshift(:pointer) # prepend *this* pointer
			}

			Class.new(FFI::Struct) {
				const_set :ParentVtbl, parent_vtbl

				const_set :Spec, {}
				self::Spec.merge!(self::ParentVtbl::Spec) if self::ParentVtbl
				self::Spec.merge!(spec)

				layout_args = self::Spec.map { |name, sig|
					params, ret = sig
					[name, callback(params, ret)]
				}
				layout_args.flatten!
				layout(*layout_args)
			}
		end
	end

	module COMInterface_
		def self.[](vtbl, siid)
			Class.new {
				const_set :Vtbl, vtbl
				const_set :IID, WindowsCOM::GUID[siid]

				def initialize(pointer)
					@vptr = COMVptr_.new(pointer)
					@vtbl = self.class::Vtbl.new(@vptr[:lpVtbl])
				end

				attr_reader :vptr, :vtbl

				self::Vtbl.members.each { |name, sig|
					define_method(name) { |*args|
						hresult = @vtbl[name].call(@vptr, *args)

						raise "#{self}.#{name} failed (hresult: #{format('%#08x', hresult)})" if
							WindowsCOM::FAILED(hresult)

						hresult
					}
				}
			}
		end
	end

	module COMInterface
		def self.[](parent_iface, siid, spec)
			vtbl = COMVtbl_[(parent_iface) ? parent_iface::Vtbl : nil, spec]

			COMInterface_[vtbl, siid]
		end
	end

	module COMFactory
		def self.[](iface, sclsid)
			Class.new(iface) {
				const_set :CLSID, WindowsCOM::GUID[sclsid]

				def initialize(clsctx = CLSCTX_INPROC)
					FFI::MemoryPointer.new(:pointer) { |ppv|
						DetonateHresult(:CoCreateInstance,
							self.class::CLSID, nil, clsctx, self.class::IID, ppv
						)

						super(ppv.read_pointer)
					}
				end
			}
		end
	end
end
