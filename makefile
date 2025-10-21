
all: openvbs.dll sample.exe sample-regfree.exe
	@echo DONE

openvbs.dll:
	cl /Ox /EHa /D_CRT_STDIO_ISO_WIDE_SPECIFIERS openvbs.cpp jujube.cpp /link /OPT:NOICF /DLL /OUT:openvbs.dll /DEF:openvbs.def advapi32.lib Ole32.lib
	del *.obj *.lib *.exp

sample.exe:
	csc sample.cs

sample-regfree.exe:
	csc sample-regfree.cs

clean:
	del openvbs.dll sample.exe sample-regfree.exe
