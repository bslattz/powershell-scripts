param (
    [string]$target = "../bash"
)

. .\Utils.ps1
. .\GitUtils.ps1
. .\GitPrompt.ps1

$CDR = $PWD

cd $target

Write-Output $target

$s = Get-GitStatus $target

Write-Output ConvertTo-Json $s

$s = Write-GitStatus $s

Write-Output $s

cd $CDR

