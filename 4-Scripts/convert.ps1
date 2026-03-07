param([string]$Action = "toL5X")
$acdPath = "..\1-Production\TankControl.acd"
$l5xPath = "..\1-Production\TankControl.l5x"
 
if ($Action -eq "toL5X") {
    plc-convert --input $acdPath --output $l5xPath
    Write-Host "ACD已转换为L5X" -ForegroundColor Green
} elseif ($Action -eq "toACD") {
    plc-convert --input $l5xPath --output $acdPath
    Write-Host "L5X已转换回ACD" -ForegroundColor Green
}