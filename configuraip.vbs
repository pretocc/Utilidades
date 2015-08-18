'Configure the active network adapter
'Configura el adaptador de red activo

strcomputer = "."
 Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

 Set colItems = objWMIService.ExecQuery _
 ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

 strCount = 1

 For Each objitem in colitems
 If strCount = 1 Then
 strIPAddress = Join(objitem.IPAddress, ",")
 IP = stripaddress
 strCount = strCount + 1
 Else
 End If

 next

 Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 Set colNetAdapters = objWMIService.ExecQuery _
 ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
 strIPAddress = Array("192.168.0.100")
 strSubnetMask = Array("255.255.255.0")
 strGateway = Array("192.168.0.1")
 strDNSServers = Array("8.8.8.8","192.168.0.1")
 For Each objNetAdapter in colNetAdapters
 errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
 errGateways = objNetAdapter.SetGateways(strGateway)
 objNetAdapter.SetDNSServerSearchOrder(strDNSServers)
 Next