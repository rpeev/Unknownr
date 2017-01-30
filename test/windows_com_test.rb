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
    guid1 = GUIDFromString('00000000-0000-0000-0000-000000000001')
    guid2 = GUIDFromString('00000000-0000-0000-0000-000000000001')
    guid3 = GUIDFromString('00000000-0000-0000-0000-000000000002')

    assert GUIDEqual(guid1, guid2)
    refute GUIDEqual(guid2, guid3)
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
    assert GUIDEqual(IFoo::IID, GUIDFromString('00000000-0000-0000-0000-000000000001'))
    assert IFoo::Vtbl::Spec == {Meth1: [[:pointer], :long]}
    assert IFoo::Vtbl.members == [:Meth1]

    assert IBar::Vtbl::ParentVtbl == IFoo::Vtbl
    assert GUIDEqual(IBar::IID, GUIDFromString('00000000-0000-0000-0000-000000000002'))
    assert IBar::Vtbl::Spec == {Meth1: [[:pointer], :long], Meth2: [[:pointer], :long]}
    assert IBar::Vtbl.members == [:Meth1, :Meth2]
  end
end
