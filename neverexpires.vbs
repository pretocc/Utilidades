'Tick the don't expire option in the user properties.
'Marca la casilla de que la contraseña nunca expira en las propiedades del usuario

Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
 
strDomainOrWorkgroup = "dominio.es"
strComputer = "localhost"
strUser = "local_user"
 
Set objUser = GetObject("WinNT://" & strDomainOrWorkgroup & "/" & _
    strComputer & "/" & strUser & ",User")
 
objUserFlags = objUser.Get("UserFlags")
objPasswordExpirationFlag = objUserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", objPasswordExpirationFlag 
objUser.SetInfo