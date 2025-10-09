
all: openvbs.dll sample.exe
	@echo DONE

openvbs.dll:
	cl /Ox /EHa /D_CRT_STDIO_ISO_WIDE_SPECIFIERS openvbs.cpp jujube.cpp /link /DLL /OUT:openvbs.dll /DEF:openvbs.def advapi32.lib Ole32.lib
	del *.obj *.lib *.exp

sample.exe:
	csc sample.cs

clean:
	del openvbs.dll sample.exe
