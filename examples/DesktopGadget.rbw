require 'windows_com'

include WindowsCOM

IDesktopGadget = COMInterface[IUnknown,
	'c1646bc4-f298-4f91-a204-eb2dd1709d1a',

	RunGadget: [[:buffer_in], :long]
]

DesktopGadget = COMFactory[IDesktopGadget, '924ccc1b-6562-4c85-8657-d177925222b6']

UsingCOMObjects(DesktopGadget.new) { |dg|
	dg.RunGadget(
		"#{ENV['ProgramFiles']}\\Windows Sidebar\\Gadgets\\Clock.Gadget\0".encode!('utf-16le')
	)
}
