# Microsoft Developer Studio Generated NMAKE File, Based on IbOleDb.dsp
!IF "$(CFG)" == ""
CFG=IbOleDb - Win32 Debug
!MESSAGE No configuration specified. Defaulting to IbOleDb - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "IbOleDb - Win32 Release" && "$(CFG)" != "IbOleDb - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "IbOleDb.mak" CFG="IbOleDb - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "IbOleDb - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "IbOleDb - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 

!IF  "$(CFG)" == "IbOleDb - Win32 Release"

OUTDIR=.\x86\Release
INTDIR=.\x86\Release
# Begin Custom Macros
OutDir=.\x86\Release
# End Custom Macros

ALL : "$(OUTDIR)\IbOleDb.dll"


CLEAN :
	-@erase "$(INTDIR)\blob.obj"
	-@erase "$(INTDIR)\command.obj"
	-@erase "$(INTDIR)\convert.obj"
	-@erase "$(INTDIR)\dso.obj"
	-@erase "$(INTDIR)\ibaccessor.obj"
	-@erase "$(INTDIR)\iberror.obj"
	-@erase "$(INTDIR)\iboledb.obj"
	-@erase "$(INTDIR)\property.obj"
	-@erase "$(INTDIR)\rowset.obj"
	-@erase "$(INTDIR)\schemarowset.obj"
	-@erase "$(INTDIR)\session.obj"
	-@erase "$(INTDIR)\util.obj"
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(OUTDIR)\IbOleDb.dll"
	-@erase "$(OUTDIR)\IbOleDb.exp"
	-@erase "$(OUTDIR)\IbOleDb.lib"
	-@erase ".\x86\Release\iboledb.res"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MT /W3 /GX /O2 /I "C:\Program Files\InterBase Corp\InterBase\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "IBOLEDB_EXPORTS" /Fp"$(INTDIR)\IbOleDb.pch" /YX /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

.c{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/nologo /D "NDEBUG" /mktyplib203 /win32 
RSC=rc.exe
RSC_PROJ=/l 0x409 /fo"$(INTDIR)\iboledb.res" /d "NDEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\IbOleDb.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=fbclient_ms.lib /nologo /dll /incremental:no /pdb:"$(OUTDIR)\IbOleDb.pdb" /machine:I386 /def:".\src\iboledb.def" /out:"$(OUTDIR)\IbOleDb.dll" /implib:"$(OUTDIR)\IbOleDb.lib" /libpath:"c:\program files\interbase corp\interbase\lib" 
DEF_FILE= \
	".\src\iboledb.def"
LINK32_OBJS= \
	"$(INTDIR)\blob.obj" \
	"$(INTDIR)\command.obj" \
	"$(INTDIR)\convert.obj" \
	"$(INTDIR)\dso.obj" \
	"$(INTDIR)\ibaccessor.obj" \
	"$(INTDIR)\iberror.obj" \
	"$(INTDIR)\iboledb.obj" \
	"$(INTDIR)\property.obj" \
	"$(INTDIR)\rowset.obj" \
	"$(INTDIR)\schemarowset.obj" \
	"$(INTDIR)\session.obj" \
	"$(INTDIR)\util.obj" \
	".\x86\Release\iboledb.res"

"$(OUTDIR)\IbOleDb.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"

OUTDIR=.\x86\Debug
INTDIR=.\x86\Debug
# Begin Custom Macros
OutDir=.\x86\Debug
# End Custom Macros

ALL : "$(OUTDIR)\IbOleDb.dll" "$(OUTDIR)\IbOleDb.bsc"


CLEAN :
	-@erase "$(INTDIR)\blob.obj"
	-@erase "$(INTDIR)\blob.sbr"
	-@erase "$(INTDIR)\command.obj"
	-@erase "$(INTDIR)\command.sbr"
	-@erase "$(INTDIR)\convert.obj"
	-@erase "$(INTDIR)\convert.sbr"
	-@erase "$(INTDIR)\dso.obj"
	-@erase "$(INTDIR)\dso.sbr"
	-@erase "$(INTDIR)\ibaccessor.obj"
	-@erase "$(INTDIR)\ibaccessor.sbr"
	-@erase "$(INTDIR)\iberror.obj"
	-@erase "$(INTDIR)\iberror.sbr"
	-@erase "$(INTDIR)\iboledb.obj"
	-@erase "$(INTDIR)\iboledb.sbr"
	-@erase "$(INTDIR)\property.obj"
	-@erase "$(INTDIR)\property.sbr"
	-@erase "$(INTDIR)\rowset.obj"
	-@erase "$(INTDIR)\rowset.sbr"
	-@erase "$(INTDIR)\schemarowset.obj"
	-@erase "$(INTDIR)\schemarowset.sbr"
	-@erase "$(INTDIR)\session.obj"
	-@erase "$(INTDIR)\session.sbr"
	-@erase "$(INTDIR)\util.obj"
	-@erase "$(INTDIR)\util.sbr"
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(INTDIR)\vc60.pdb"
	-@erase "$(OUTDIR)\IbOleDb.bsc"
	-@erase "$(OUTDIR)\IbOleDb.dll"
	-@erase "$(OUTDIR)\IbOleDb.exp"
	-@erase "$(OUTDIR)\IbOleDb.ilk"
	-@erase "$(OUTDIR)\IbOleDb.lib"
	-@erase "$(OUTDIR)\IbOleDb.pdb"
	-@erase ".\x86\Debug\iboledb.res"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MTd /W3 /Gm /GX /ZI /Od /I "C:\Program Files\InterBase Corp\InterBase\include" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "IBOLEDB_EXPORTS" /FR"$(INTDIR)\\" /Fp"$(INTDIR)\IbOleDb.pch" /YX /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /GZ /c 

.c{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/nologo /D "_DEBUG" /mktyplib203 /win32 
RSC=rc.exe
RSC_PROJ=/l 0x409 /fo"$(INTDIR)\iboledb.res" /d "_DEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\IbOleDb.bsc" 
BSC32_SBRS= \
	"$(INTDIR)\blob.sbr" \
	"$(INTDIR)\command.sbr" \
	"$(INTDIR)\convert.sbr" \
	"$(INTDIR)\dso.sbr" \
	"$(INTDIR)\ibaccessor.sbr" \
	"$(INTDIR)\iberror.sbr" \
	"$(INTDIR)\iboledb.sbr" \
	"$(INTDIR)\property.sbr" \
	"$(INTDIR)\rowset.sbr" \
	"$(INTDIR)\schemarowset.sbr" \
	"$(INTDIR)\session.sbr" \
	"$(INTDIR)\util.sbr"

"$(OUTDIR)\IbOleDb.bsc" : "$(OUTDIR)" $(BSC32_SBRS)
    $(BSC32) @<<
  $(BSC32_FLAGS) $(BSC32_SBRS)
<<

LINK32=link.exe
LINK32_FLAGS=fbclient_ms.lib /nologo /dll /incremental:yes /pdb:"$(OUTDIR)\IbOleDb.pdb" /debug /machine:I386 /def:".\src\iboledb.def" /out:"$(OUTDIR)\IbOleDb.dll" /implib:"$(OUTDIR)\IbOleDb.lib" /pdbtype:sept /libpath:"c:\program files\interbase corp\interbase\lib" 
DEF_FILE= \
	".\src\iboledb.def"
LINK32_OBJS= \
	"$(INTDIR)\blob.obj" \
	"$(INTDIR)\command.obj" \
	"$(INTDIR)\convert.obj" \
	"$(INTDIR)\dso.obj" \
	"$(INTDIR)\ibaccessor.obj" \
	"$(INTDIR)\iberror.obj" \
	"$(INTDIR)\iboledb.obj" \
	"$(INTDIR)\property.obj" \
	"$(INTDIR)\rowset.obj" \
	"$(INTDIR)\schemarowset.obj" \
	"$(INTDIR)\session.obj" \
	"$(INTDIR)\util.obj" \
	".\x86\Debug\iboledb.res"

"$(OUTDIR)\IbOleDb.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ENDIF 


!IF "$(NO_EXTERNAL_DEPS)" != "1"
!IF EXISTS("IbOleDb.dep")
!INCLUDE "IbOleDb.dep"
!ELSE 
!MESSAGE Warning: cannot find "IbOleDb.dep"
!ENDIF 
!ENDIF 


!IF "$(CFG)" == "IbOleDb - Win32 Release" || "$(CFG)" == "IbOleDb - Win32 Debug"
SOURCE=.\src\blob.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\blob.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\blob.obj"	"$(INTDIR)\blob.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\command.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\command.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\command.obj"	"$(INTDIR)\command.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\convert.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\convert.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\convert.obj"	"$(INTDIR)\convert.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\dso.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\dso.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\dso.obj"	"$(INTDIR)\dso.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\ibaccessor.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\ibaccessor.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\ibaccessor.obj"	"$(INTDIR)\ibaccessor.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\iberror.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\iberror.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\iberror.obj"	"$(INTDIR)\iberror.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\iboledb.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\iboledb.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\iboledb.obj"	"$(INTDIR)\iboledb.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\property.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\property.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\property.obj"	"$(INTDIR)\property.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\rowset.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\rowset.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\rowset.obj"	"$(INTDIR)\rowset.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\schemarowset.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\schemarowset.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\schemarowset.obj"	"$(INTDIR)\schemarowset.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\session.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\session.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\session.obj"	"$(INTDIR)\session.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\src\util.cpp

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


"$(INTDIR)\util.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


"$(INTDIR)\util.obj"	"$(INTDIR)\util.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\res\iboledb.rc

!IF  "$(CFG)" == "IbOleDb - Win32 Release"


".\x86\Release\iboledb.res" : $(SOURCE)
	$(RSC) /l 0x409 /fo"x86\Release/iboledb.res" /i "res" /d "NDEBUG" $(SOURCE)


!ELSEIF  "$(CFG)" == "IbOleDb - Win32 Debug"


".\x86\Debug\iboledb.res" : $(SOURCE)
	$(RSC) /l 0x409 /fo"x86\Debug/iboledb.res" /i "res" /d "_DEBUG" $(SOURCE)


!ENDIF 


!ENDIF 

