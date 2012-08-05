 Dim sOriginFolder, sDestinationFolder, oFSO, install
 Set oFSO = CreateObject("Scripting.FileSystemObject") 
 Set oShell = CreateObject( "WScript.Shell" ) 
 install = 1
 sDestinationFolder = oShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Excel\XLSTART"
 Files = array("num2curr.xla")


 'See if excel is installed
 Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
 oReg.GetStringValue &H80000000,"Excel.Application","",dwValue
 if IsNull(dwValue) then
     wscript.echo("Any version of Excel is not installed")
     install = 0
 end if

 'See if Excel is running
 On Error Resume Next 
 Dim excel: Set excel = GetObject(, "Excel.Application") 
 If Err.Number = 0 Then 
     wscript.echo("Excel is running,Please close Excel and try installing again")
     install=0
 End If 
 Err.Clear
 On Error goto 0

'main installation process
 for each sFile in Files
     If Not oFSO.FileExists(sDestinationFolder & "\" & oFSO.GetFileName(sFile)) and install=1 Then 
        oFSO.GetFile(sFile).Copy sDestinationFolder & "\" & oFSO.GetFileName(sFile),True 
     elseif install=-1 then
        if oFSO.FileExists(sDestinationFolder & "\" & oFSO.GetFileName(sFile)) then
             oFSO.DeleteFile sDestinationFolder & "\" & oFSO.GetFileName(sFile)
        end if
     elseif install=1 then
         if msgbox("Program already installed do you want to uninstall ?",4) = 6 then
             install = -1
             oFSO.DeleteFile sDestinationFolder & "\" & oFSO.GetFileName(sFile)    
         else
             install = 0 
             exit for
         end if
     End If 
 next

 if install = -1 then
     wscript.echo("Uninstallation Completed")
 elseif install = 1 then
     wscript.echo("Installation Completed")
 else
     wscript.echo("Installation aborted")
 end if
 
