echo off

REM Copy the assemblies needed for a Project Server event handler, 
REM from the Project Server computer to a directory share on a development computer.
REM Run as administrator.

set SHARE=\\DEV_PC\ShareName

set SCHEMA=C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Project.Schema\14.0.0.0__71e9bce111e9429c
set PROJ_SERVER_LIBS="c:\Program Files\Microsoft Office Servers\14.0\Bin"
set SHAREPOINT_LIBS="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\14\ISAPI"

xcopy /y %SCHEMA%\*.dll %SHARE%
xcopy /y %PROJ_SERVER_LIBS%\Microsoft.Office.Project.Server.Events.Receivers.dll %SHARE%
xcopy /y %PROJ_SERVER_LIBS%\Microsoft.Office.Project.Server.Library.dll %SHARE%
xcopy /y %SHAREPOINT_LIBS%\Microsoft.SharePoint.dll %SHARE%