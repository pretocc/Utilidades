'Change the name of the network adapter
'Cambia el nombre del adaptador de red.

Const NETWORK_CONNECTIONS = &H31&
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(NETWORK_CONNECTIONS)
Set colItems = objFolder.Items
For Each objItem in colItems
   If objItem.Name = "Conexión de área local" Then
      objItem.Name = "Red"
   End If
  
Next
