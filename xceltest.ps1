$work1 = Import-Excel -Path 'D:\powershell\Docs\Book1.xlsx' -WorksheetName 'PilotShip'
$work2 = Import-Excel -Path 'D:\powershell\Docs\Book1.xlsx' -WorksheetName 'JobPlat'
foreach ( $row in $work1) {
    $idnow = $row.ID
    Write-Host $row.ID
    Write-Host "checking"
    if ( $row.ID -contains $work2.ID) {
    Write-host "Booya $idnow"
    } Else {write-host "booNoooooooooo $idnow"}
    }



