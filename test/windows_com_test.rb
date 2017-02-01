require 'minitest'
require 'minitest/autorun'

require_relative '../lib/windows_com'

include WindowsCOM

class WindowsCOMTest < Minitest::Test
  def test_WINDOWS_COM_xxx
    assert_match %r{^\d+\.\d+\.\d+(\.\d+)?$}, WINDOWS_COM_VERSION
    assert WINDOWS_COM_OLE_INIT
  end

  def test_GUID
    guid1 = GUID['00000000-0000-0000-0000-000000000001']
    guid2 = GUID['00000000-0000-0000-0000-000000000001']
    guid3 = GUID['00000000-0000-0000-0000-000000000002']

    assert guid1 == guid2
    refute guid1 != guid2
    assert guid2 != guid3
    refute guid2 == guid3
  end

  def test_PROPERTYKEY
    propkey1 = PROPERTYKEY[VT_BOOL, 1]
    propkey2 = PROPERTYKEY[VT_BOOL, 1]
    propkey3 = PROPERTYKEY[VT_BOOL, 2]

    assert propkey1 == propkey2
    refute propkey1 != propkey2
    assert propkey2 != propkey3
    refute propkey2 == propkey3
  end

  def test_FFIStructAnonymousAccess
    v = VARIANT.new

    # direct member access
    assert v[:decVal].class == DECIMAL

    # level one _ member access
    v[:vt] = VT_BOOL
    assert v[:vt] == VT_BOOL

    # level two _ member access
    v[:intVal] = 42
    assert v[:intVal] == 42

    # nonexisting field access
    assert_raises(ArgumentError) {
      v[:foo]
    }
  end

  IFoo = COMInterface[nil,
    '00000000-0000-0000-0000-000000000001',

    Meth1: [[], :long]
  ]

  IBar = COMInterface[IFoo,
    '00000000-0000-0000-0000-000000000002',

    Meth2: [[], :long]
  ]

  def test_COMInterface
    assert_nil IFoo::Vtbl::ParentVtbl
    assert IFoo::IID == GUID['00000000-0000-0000-0000-000000000001']
    assert IFoo::Vtbl::Spec == {Meth1: [[:pointer], :long]}
    assert IFoo::Vtbl.members == [:Meth1]

    assert IBar::Vtbl::ParentVtbl == IFoo::Vtbl
    assert IBar::IID == GUID['00000000-0000-0000-0000-000000000002']
    assert IBar::Vtbl::Spec == {Meth1: [[:pointer], :long], Meth2: [[:pointer], :long]}
    assert IBar::Vtbl.members == [:Meth1, :Meth2]
  end
end
