require 'unknownr'

include Unknownr::Windows

InitializeOle()

IDesktopGadget = COM::Interface[IUnknown,
	GUID['c1646bc4-f298-4f91-a204-eb2dd1709d1a'],

	RunGadget: [[:buffer_in], :long]
]

DesktopGadget = COM::Factory[IDesktopGadget, GUID['924ccc1b-6562-4c85-8657-d177925222b6']]

dg = DesktopGadget.new

begin
	dg.RunGadget(
		"#{ENV['ProgramFiles']}/Windows Sidebar/Gadgets/Clock.Gadget\0".
			gsub('/', '\\').
			encode('utf-16le')
	)
ensure
	dg.Release
end
