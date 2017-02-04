# Release Notes

## 2.1.1

Sync important example with changes in related gem (windows_gui)

## 2.1.0

`COMVtbl_` doesn't prepend `this` pointers to vtbl specs

## 2.0.2

Implement the ability to trace COM call arguments (set `WINDOWS_COM_TRACE_CALL_ARGS`)

## 2.0.1

Allow `COMInterface_` instances to be passed to code expecting pointers without calling `#vptr`

## 2.0.0

- Add `COMCallback` module for __implementing__ COM interfaces in Ruby
- enhance some bound FFI structs with useful methods
- improve code

## 1.0.0

Rename library to windows_com and ensure it works with recent ruby

## 0.2.2

Recover source from gem
