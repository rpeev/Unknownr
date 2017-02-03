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
# uia is an instance of UIA, a child class of IUIApplicationImpl
uif.Initialize(hwnd, uia)
```

## Install

gem install windows_com

## Use

See examples folder (the UIRibbon example requires windows_gui gem)
