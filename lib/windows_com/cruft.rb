__END__
module COMHelpers
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

module COMCallback
  def self.[](iface)
    Class.new(FFI::Struct) {
      send(:include, COMHelpers)

      layout \
        :lpVtbl, :pointer

      def initialize(opts = {})
        @vtbl, @refc = iface::VTBL.new, 1

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
        if [IUnknown::IID, iface::IID].any? { |iid| windows_com_memcmp(riid, iid, iid.size) == 0 }
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

      (iface::VTBL.members - IUnknown::VTBL.members).each { |name|
        define_method(name) { |*args|
          E_NOTIMPL
        }
      }
    }
  end
end

module AnonymousFFIStructSupport
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

# PROPERTYKEY
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

# PROPVARIANT
def ==(other) windows_com_memcmp(other, self, size) == 0 end

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
