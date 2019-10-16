
 Write-Host "Powershell script created by: zdmurai on 8/2/19 (v1)"
 Write-Host "`n"


Function Format-FileSize() {
Param ([int]$size)
If ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00} GB", $size / 1GB)}
ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00} MB", $size / 1MB)}
ElseIf ($size -gt 1KB) {[string]::Format("{0:0.00} kB", $size / 1KB)}
ElseIf ($size -gt 0) {[string]::Format("{0:0.00} B", $size)}
Else {""}
}

function Format-Color([hashtable] $Colors = @{}, [switch] $SimpleMatch) {
    $lines = ($input | Out-String) -replace "`r", "" -split "`n"
    foreach($line in $lines) {
        $color = ''
        foreach($pattern in $Colors.Keys){
            if(!$SimpleMatch -and $line -match $pattern) { $color = $Colors[$pattern] }
            elseif ($SimpleMatch -and $line -like $pattern) { $color = $Colors[$pattern] }
        }
        if($color) {
            Write-Host -ForegroundColor $color $line -BackgroundColor black 
        } else {
            Write-Host $line
        }
    }
}


$csvs = Get-ChildItem .\*  
$y=$csvs.Count


Write-Host "Detected the following CSV files: ($y)"
foreach ($csv in $csvs)
{
Write-Host " "$csv.Name
}

$csv = "csv"
$filename = Read-Host -Prompt 'insert filename'

if([string]::isNullOrWhiteSpace($filename)) {
    Write-Host "Filename Empty"   
    exit
}

$filename = "$filename.$csv"


#converted date to GMT + 8
$gmt = Import-Csv $filename
$gmt[-1].date = Get-Date ([System.TimeZoneInfo]::ConvertTime([datetime]($gmt[-1].date),([System.TimeZoneInfo]::FindSystemTimeZoneById('Singapore Standard Time')))) -Format 'HH:mm "GMT"+8'
$g = $gmt | select -last 1 -ExpandProperty date 


#original date
$csvOriginal = Import-Csv $filename
$t = $csvOriginal[-1].date 




$csv =Import-csv $filename 
$csv[-1].date = "$t - $g"
$csv | select -last 1 | Format-Color @{'GMT' = 'Green'}


write-host "Raw count: "$csv.count 

if($csv.count -gt 25000) {
     Write-Host "More files to download!!" -ForegroundColor Red -BackgroundColor black 
     Write-Host "`a `a" #beeps will produce
    


    

} else {
    "`n"

    # consolidate
    # remove empty
    # sort MID unique


   
    Write-Host "Consolidating..." -ForegroundColor red -BackgroundColor black 
    cat *.csv | sc all.csv

    Set-Content -Path .\removeEmpty.csv -Value (get-content -Path all.csv  | Select-String -Pattern 'empty subject' -NotMatch)
    $a = Import-Csv .\removeEmpty.csv | Sort-Object MID -Unique 
    $a | Export-csv .\sorted.csv -NoTypeInformation
    #Get-ChildItem -Filter *.csv | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv .\ha.csv    -NoTypeInformation -Append
    
    
    write-host "..Deleting all.csv"
    Remove-Item "all.csv"
    write-host "..Deleting removeEmpty.csv"
    Remove-Item "removeEmpty.csv"
    #$a = ".\all.csv"
    Write-Host "..Done!" -ForegroundColor green -BackgroundColor blue 



    #cat *.csv | sc "all.csv" #merge all .csv
    #Import-csv "all.csv" | Sort-Object -Unique MID | Export-Csv "test.csv" -NoTypeInformation

    #$size=Format-FileSize((Get-Item $a ).length)
    #Write-Host "created: $a @ $size"
    #Write-Host "`a `a" #>

       
}

