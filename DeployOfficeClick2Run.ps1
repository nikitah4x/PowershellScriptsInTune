##This script will download Office Click2Run installatin using 'wget'##
##This script is using file and registry key detection method##

$FileName = "C:\Program Files\Microsoft Office"
$RegKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365BusinessRetail - en-us"
if ((Test-Path $Regkey) -Or (Test-Path $FileName)) {
  Write-Host "Microsoft Office is already installed on this device!"
  exit
}else {
    wget "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=Def&language=en-us&TaxRegion=db&correlationId=8add0e0f-36e5-4b52-94d2-bdd0dc50304e&token=019ba733-31b5-4695-865b-5efa2a193473&version=O16GA&source=O15OLSO365&Br=2" -OutFile "setup.exe"
    Start-Sleep -s 10
    Copy-Item "setup.exe" -Destination "C:\"
    Start-Sleep -s 5
    Remove-Item "setup.exe"
    Start-Sleep -s 5
    cd "C:\"
    Start-Sleep -s 5
    Write-Host "DONE!"
    .\setup.exe  
}