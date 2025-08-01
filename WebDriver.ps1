clear

$Path = "D:\PowerDMS"

if (Test-Path "$Path\EdgeVersion.txt")
{
}
else
{
  "0.0.0.0" | Out-File -FilePath "$Path\EdgeVersion.txt"
}

$PreviousVersion = Get-Content -Path "$Path\EdgeVersion.txt"

Write-Output "Previous Version = $PreviousVersion"
Write-Output ""

$InstalledVersion = (Get-ItemProperty -Path HKCU:\Software\Microsoft\Edge\BLBeacon -Name version).version

Write-Output "Installed Version = $InstalledVersion"
Write-Output ""

If ($PreviousVersion -eq $InstalledVersion)
{
  Write-Output 'Web Driver does need to be updated.'
}
Else
{
  Write-Output 'Web Driver will be updated.'
  Write-Output ""

  if (Test-Path "$Path\msedgedriver.exe")
  {
    Remove-Item "$Path\msedgedriver.exe" -Force
  }

  [Net.ServicePointManager]::SecurityProtocol =
  [Net.SecurityProtocolType]::Tls12

  $WebClient = New-Object System.Net.WebClient

  #$URL = "https://msedgedriver.azureedge.net/$InstalledVersion/edgedriver_win64.zip"
  $URL = "https://msedgedriver.microsoft.com/$InstalledVersion/edgedriver_win64.zip"

  Write-Output $URL
  
  $ZipFile = "$Path\edgedriver_win64.zip"

  $WebClient.DownloadFile($URL, $ZipFile)

  Expand-Archive $ZipFile -DestinationPath $Path -Force

  if (Test-Path $ZipFile)
  {
    Remove-Item $ZipFile -Force
  }

  if (Test-Path "$Path\Driver_Notes")
  {
    Remove-Item "$Path\Driver_Notes" -Force -Recurse
  }

  $InstalledVersion | Out-File -FilePath "$Path\EdgeVersion.txt"
}

Write-Output ""