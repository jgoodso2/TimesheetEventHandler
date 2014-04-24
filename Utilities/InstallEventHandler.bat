echo off

@SET EVENT_HANDLERS="C:\Program Files\Microsoft Office Servers\14.0\Bin\ProjectServerEventHandlers"

REM To deploy to a production server, copy the event handler to the ProjectServerEventHandlers subdirectory.
REM xcopy /y ..\bin\debug\TestCreatingProject.dll %EVENT_HANDLERS%
REM xcopy /y ..\bin\debug\TestCreatingProject.pdb %EVENT_HANDLERS%

REM To debug the event handler, register it in the GAC.
gacutil.exe /u TestCreatingProject
gacutil.exe /i ..\bin\debug\TestCreatingProject.dll

REM pause