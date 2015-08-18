# Insert the computer objects into AD using a CSV list
# Inserta las cuentas de equipo que contiene el CSV en el AD

Import-Module ActiveDirectory

$CSV='NOMBRE_PCS.csv'
$OU='OU=Contabilidad,OU=Clientes,OU=Empresa,DC=dominio,DC=local'
$PCS = Import-Csv -Path $CSV

ForEach ($Item in $PCS) {
 Write-Host $Item.Computer
 New-ADComputer -Name $Item.Computer -Path $OU -Enabled $True 
 }
