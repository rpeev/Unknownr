# windows_com

Ruby FFI (x86) bindings to essential COM related Windows APIs

![Screenshot](./screenshot.png)

## Features

- convenient DSL for binding COM interface definitions

```ruby
IUIApplication = COMInterface[IUnknown,
  'D428903C-729A-491d-910D-682A08FF2522',

  OnViewChanged: [[:uint, :int, :pointer, :int, :int], :long],
  OnCreateUICommand: [[:uint, :int, :pointer], :long],
  OnDestroyUICommand: [[:uint, :int, :pointer], :long]
]

IUIApplicationImpl = COMCallback[IUIApplication]
```

- straightforward usage of the resulting Ruby classes

```ruby
class UIA < IUIApplicationImpl
  def initialize(uich)
    @uich = uich

    super() # wire COM stuff
  end

  attr_reader :uich

  # COM interface method implementations

  def OnCreateUICommand(*args)
    uich.QueryInterface(uich.class::IID, args[-1])

    S_OK
  end
end
```

- transparent interop with code expecting COM interface pointers

```ruby
uif.Initialize(hwnd, uia)
```

## Conventions

Classes starting with capital __I__ are wrappers around raw COM interface pointers inheriting from `COMInterface_` (e.g. `IUIApplication`). They contain `COMVptr_` and `COMVtbl_` (`FFI::Struct` implementations) instances wired to the corresponding COM interface implementation function pointers and are (usually) not directly instantiated by application code.

Classes __not__ starting with capital __I__ are COM factories (obtain COM interface implementations using `CoCreateInstance`) inheriting from the corresponding COM interface class (e.g. `UIFramework` is a factory for creating instances of `IUIFramework`). They are directly instantiated by application code and create system COM objects implementing the corresponding interface.

Classes starting with capital __I__ and ending with __Impl__ are COM callbacks inheriting from the corresponding COM interface class (e.g. `IUIApplicationImpl` is a base class for implementing the `IUIApplication` interface in Ruby). They are usually subclassed, implement (some) of the corresponding COM interface methods and then instantiated by application code.

All kinds of classes can be freely subclassed and used regardles of their COM duties (just don't forget to call __super__ appropriately in the subclass `#initialize` method). `COMInterface_` (their common base) implements `#to_ptr` (returning the corresponding `COMVptr_` instance pointer), so the instances can be directly passed to code expecting COM interface pointers. In a scenario where the COM stuff is used as instance variables of some other class, but it is desirable for the class instances to be passed transparently to code expecting COM interface pointers, just define #to_ptr calling the appropriate COM object `#to_ptr`.

## Install

gem install windows_com

## Use

See examples folder (the UIRibbon example requires windows_gui gem)
