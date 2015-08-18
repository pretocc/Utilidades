'=========================================================================='
' Title: List Printers.vbs' 
' Date: 02/23/2010' 
' Author: Bradley Buskey' 
' Version: 1.00' 
' Updated: 02/23/2010' 
' Purpose: List all printers attached to a workstation' 
'=========================================================================='
'=========================================================================='
' Date 03/29/2010
' updated: Chris Daws
' to include mapped network printers
'==========================================================================
'Date 06/08/2014
'updated: Santiago Prego
'to correct an error when no user is logged in the system
'==========================================================================
Const ForAppending = 8 
Const ForReading = 1 

Dim WshNetwork, objPrinter, intDrive, intNetLetter

strComputer = inputbox("Please enter the computer name or IP address.","Computer Name",".") 
Set WshNetwork = CreateObject("WScript.Network") 
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer") 
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48) 
Set WshShell = WScript.CreateObject("WScript.Shell") 
Set objFSO = CreateObject("Scripting.FileSystemObject") 

For Each objItem in colItems 
UserName = objItem.UserName 

'Se comprueba que haya una sesión iniciada por algún usuario.
If IsNull(UserName) Then
MsgBox "No user logged in the system."
Wscript.Quit
End If

arrUserName = Split(UserName, "\", -1, 1) 
varUserName = arrUserName(1) 
Next 

filOutput = varUserName & ".txt" 

If objFSO.FileExists(filOutput) Then 
objFSO.DeleteFile(filOutput) 
End If 

Set objOutputFile = objFSO.OpenTextFile (filOutput, ForAppending, True) 
For Each objPrinter in colInstalledPrinters 
strTest = Left(objPrinter.Name, 2) 
objOutputFile.WriteLine(objPrinter.Name) 
Next 
'objOutputFile.Close


'added
Set objPrinter = WshNetwork.EnumPrinterConnections
'Set objOutputFile = objFSO.OpenTextFile (filOutput, ForAppending, True) 
If objPrinter.Count = 0 Then
WScript.Echo "No Printers Mapped "
else
For intDrive = 0 To (objPrinter.Count -1) Step 2
intNetLetter = IntNetLetter +1
printer = "UNC Path " & objPrinter.Item(intDrive) & " = " & objPrinter.Item(intDrive +1) & " Printer : " & intDrive
objOutputFile.WriteLine(printer)
Next
end if
objOutputFile.Close
'added

 

varOpen = MsgBox("Do you want to view the printers?",36,"View File?") 
If varOpen = vbYes Then 
varCommand = "notepad " & filOutput 
WshShell.Run varCommand,1,False 
End If 

Wscript.Sleep 1500 
MsgBox "Printer mappings have been stored in '" & filOutput & "'.", 64, "Script Complete" 
Wscript.Quit