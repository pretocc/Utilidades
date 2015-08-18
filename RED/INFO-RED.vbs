'########################################################################
'# This vbScript change the IP address, subnet mask and DNS servers.	#
'# Save all the info in a database, such as OLDIP, NEWIP, MAC, etc.		#
'#																		#
'# Author: Santiago Prego												#
'# 2013																	#
'########################################################################

Option Explicit

'########## VARIABLES DEFINITION ########

Dim strComputer
Dim strNewComputer 
Dim strDomainUser  
Dim strDomainPasswd
Dim strLocalUser  
Dim strLocalPasswd
Dim strIPS 
Dim strIP2 
Dim strMAC 
Dim strOLDIP 
Dim strModel 
Dim strOLDID 
Dim strNEWIP 
Dim strSTATUS 
Dim strSITUATION
Dim CurrentName
Dim database
Dim WshNetwork
Dim objWMILocator
Dim objFSO
Dim objTextFile
Dim objTextFile1
Dim strNextLine
Dim arrHosts
Dim objWMIComputer
Dim objWMIComputerSystem
Dim objWMIService
Dim colItems
Dim objItem
Dim objAddress
Dim rc
Dim strDATE
Dim arrNewDNSServerSearchOrder
Dim strDNSServer
Dim intSetDNSServers
Dim intSetIP
Dim arrNEWIP
Dim arrMASK


'########## VARIABLES DEFINITION END ######

'########## SCRIPT CONFIGURATION ##########

arrNewDNSServerSearchOrder = Array("8.8.8.8", "192.168.0.1") ' DNS Servers
strComputer     = "REMOTO-PC"								 ' Hostname of the remote computer
strDomainUser   = "user@domain.local"						 ' Domain user login name
strDomainPasswd = "password"
strLocalUser    = "local_user"								 ' Local user with admin rights
strLocalPasswd  = "password"
strDATE = Now()
strIPS = "192"	'Contain the first numbers of our subnet
strIP2 = "IP"	'Contain the five first numers of the Ip Address
strMAC = "."	'Contain the Remote MAC address
strOLDIP = "."	'Contain the old IP adrress
strModel = "."	'Contain the model of the remote network adapter
strOLDID = "."	'Contain the old name of the remote computer
strNEWIP = "192.168.0.100"		'Contain the new IP of the remote computer
arrMASK = Array("255.255.0.0")	'Contain the subnet mask
arrNEWIP = Array(strNEWIP)		'Contain the new IP of the remote computer
strSTATUS = "."		'Contain the status of the remote computer. OK = Computer successfully renamed ; ERROR = Rename failded with error ; OFFLINE = We don't have a ping response
strSITUATION = "."	'Contain some letters in the old name of the remote computer
Set database = CreateObject("ADODB.Connection") 'Contain the database object
database.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=inventarioRED.mdb" 'Open the database
Set WshNetwork = CreateObject("WScript.Network") 'Contain the Network wsh object
CurrentName = WshNetwork.ComputerName 'Returns the name of the local computer
Const ForReading = 1
Const ForAppending = 8

'############ END CONFIGURATION ###########

'############################################################
' In this function we ping the computer for check the status
'############################################################


Function Ping(strComputer)

    dim objPing, objRetStatus

    set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
      ("select * from Win32_PingStatus where address = '" & strComputer & "'")

    for each objRetStatus in objPing
        if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode<>0 then
			Ping = False
			else
            Ping = True
        end if
    Next
End Function


'########################################################################
' Here we get the array of computers and separate the wrong renaming
'########################################################################


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("hosts.txt", ForReading)
Set objTextFile1 = objFSO.OpenTextFile _
    ("error-hosts.txt", ForAppending, True)

Do Until objTextFile.AtEndOfStream 'We repeat all until the array is empty
    strNextLine = objTextFile.Readline
    arrHosts = Split(strNextLine , ",")
	'Wscript.Echo "Server name: " & arrHosts(0)
	strComputer = arrHosts(0)


'###########################
' Connect to Computer
'###########################

strOLDID = strComputer
strSITUATION = Mid(strOLDID,6,3)


if Ping(strComputer) = True Then ' Here we check if the computer is on-line

set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
objWMILocator.Security_.AuthenticationLevel = 6

On Error Resume Next

set objWMIComputer = objWMILocator.ConnectServer(strComputer,  _
           		                         "root\cimv2", _
                                                  strLocalUser, _
                                                  strLocalPasswd)


        If Err.Number <> 0 Then 
								' Save the status and write the line in error-hosts
							  strSTATUS = "" & Err.Description
						      objTextFile1.WriteLine(arrHosts(0) & " - " & strSTATUS)
        End If


set objWMIComputerSystem = objWMIComputer.Get( _
                               "Win32_ComputerSystem.Name='" & _
                               strComputer & "'")


'#################################################################################
'// Here we get the IP address, model of the Ethernet card and the MAC address
'#################################################################################

Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration " & _
        "Where IPEnabled = True" )

For Each objItem in colItems

    For Each objAddress in objItem.IPAddress
	strIP2=Left(objAddress,2)

	If strIP2 = strIPS Then

		strMODEL = objItem.caption
		strMAC = objItem.macaddress
		strOLDIP = objAddress

	End IF 
    Next

Next


Else

	'Wscript.Echo "Host " & strComputer & " could not be reached" ' If the computer are off-line save the status and write the host in ERROR-HOSTS
	strSTATUS = "OFFLINE"
	objTextFile1.WriteLine(arrHosts(0) & " - " & strSTATUS)

End If


'#########################################
' Here we insert the data in the database
'#########################################

database.Execute "INSERT INTO DATA (OLDID,NEWIP,OLDIP,SITUATION,STATUS,MODEL,MAC,EXDATE) VALUES ('" & strOLDID & "', '" & strNEWIP & "', '" & strOLDIP & "', '" & strSITUATION & "', '" & strSTATUS & "', '" & strMODEL & "', '" & strMAC &  "', '" & strDATE &  "')"


strOLDID = "."
strNEWIP = "."

Loop 'We repeat all until the array is empty

database.Close
Wscript.Echo "Proceso Finalizado"