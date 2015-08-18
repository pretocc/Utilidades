'Comprobamos que el equipo tiene instalado LibreOffice y quitamos el componente Base

Const HKEY_LOCAL_MACHINE = &H80000002

Equipo = "."  

Set objSrvMens=GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & Equipo & "\root\cimv2") 

Set products = objSrvMens.ExecQuery  ("SELECT * FROM Win32_Product WHERE Name like '%libreoffice%'") 
Set features = objSrvMens.ExecQuery  ("SELECT * FROM Win32_SoftwareFeature WHERE Name = 'gm_p_Base'") 

For Each product In products 
  Wscript.echo product.Name
  
 For Each feature in Features 
   Wscript.echo feature.Name
   p_InstallState = 4
		'Información de los tipos de estado
		'1='Default'
		'2='Advertise'
		'3='Local'
		'4='Absent'
		'5='Source'
   feature.Configure(p_InstallState) 'Con esto quitamos el componente
   
 Next
Next