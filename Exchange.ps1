$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://srv-exc-01/powershell
$HeaderXLSX = 'direction','nom','prénom','mail' 
$FileExcel = Import-Excel -Path "C:\Users\sautron_di\Downloads\Liste des managers Univers.xlsx" # -Delimiter ";" -Header $HeaderXLSX

$ExportUser = @()
$ExportCSV = @()
Start-Transcript "C:\Tools\log\Liste_des_managers_univers.txt"
foreach ($FileXLSX in $FileExcel)
    {
    #$IDUser = $FileXLSX.Nom
    write-host $FileXLSX.'prénom'";"$FileXLSX.'nom'
    $ExportUser += "$FileXLSX.'Prénom'"+";"+"$FileXLSX.'nom'"
    $IDUserNom = $FileXLSX.'nom'
    $IDUserPrenom = $FileXLSX.'prénom'
    Get-ADUser -Filter {Surname -eq $IDUserNom -and GivenName -eq $IDUserPrenom} -Properties * | Select GivenName,Surname,mail,UserPrincipalName,SamAccountName
    #$SamAccountNameFull = Get-ADUser -Filter {Surname -eq $IDUserNom -and GivenName -eq $IDUserPrenom} -Properties * | Select SamAccountName
    $ExportUser += Get-ADUser -Filter {Surname -eq $IDUserNom -and GivenName -eq $IDUserPrenom} -Properties * | Select GivenName,Surname,mail,UserPrincipalName,SamAccountName
    
    }
$ExportCSV = $ExportUser
$ExportCSV | Export-Csv -Path "C:\Tools\Liste_des_managers_univers.csv" -Encoding UTF8 -NoTypeInformation -Delimiter ";"

#Get-ADUser -Filter {Surname -eq "Gronio"} -Properties * | Select GivenName,Surname,mail

#Import-PSSession $Session
Stop-Transcript