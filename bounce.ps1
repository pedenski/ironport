
#BOUNCE TO CSV
write-host "====================="
write-host "START: S BOUNCE TO CSV"
write-host "====================="
Cat *.s | sc ball.s #Consolidate all .S
(Get-Content 'ball.s') | Where-Object { $_ -match 'careserror' } | Set-Content 'ball.s' #Check per line and get only lines with carerserror
Set-Content -Path "ball.s" -Value (get-content -Path "ball.s" | Select-String -Pattern 'delayed' -NotMatch) #Remove lines with delayed 
(Get-content 'ball.s') | ForEach-Object { $_ -replace ',',':::'} | Set-Content 'ball.s' #replace all comma (,) with (:::) this avoids CSV to delimit comma and separate it per column.
import-csv ball.s | export-csv balls.csv  -NoTypeInformation #convert *.s by exporting to csv
write-host "====================="
write-host "DONE: S BOUNCE TO CSV"
write-host "====================="




#GETTING MID AND STATE FROM BOUNCE
write-host "====================================="
write-host "START: GET MID AND STATE FROM BOUNCE"
write-host "====================================="
$thispath = Get-Location
write-host "opening $thispath\balls.csv"
$excel = new-object -comobject Excel.Application
$excel.visible = $false
$workbook = $excel.workbooks.open("$thispath\balls.csv")
$worksheet = $workbook.Worksheets.Item(1)

write-host "Inserting Formula"

$colCount = $worksheet.range("A1").currentregion.rows.count #get Column count of column A

$worksheet.range("B1:B$colCount").formula = '=MID(A1,FIND("MID",A1)+4,9)' #add formula to columb B
$worksheet.range("C1:C$colCount").formula = 
'=IF(ISNUMBER(SEARCH("spam content blocked",A1)),"rejected by destination", IF(OR(ISNUMBER(SEARCH({"does not exist","is not exist","doesnt have","aborted","expired","no such user","mailbox unavailable"},A1))),"invalid email",IF(OR(ISNUMBER(SEARCH({"dns soft","dns hard"},A1))),"invalid domain",IF(ISNUMBER(SEARCH("disabled",A1)),"mailbox disabled",IF(ISNUMBER(SEARCH("quota",A1)),"mailbox full","rejected by destination")))))'

$workbook.SaveAs("$thispath\bounce.csv")

$workbook.Close($false)
$excel.Quit()

write-host "======================================="
write-host "DONE: GETTING MID AND STATE FROM BOUNCE"
write-host "========================================"


# need to release COM references and ensures no remaining references exist which will keep Excel.exe open 
# Without performing the above, if you look in Task Manager, you may see Excel still running...in some cases, many copies.
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Remove-Variable -Name excel


#ADD COLUMN DATE MID STATE

write-host "====================================="
write-host "START: ADD COLUMN DATE / MID / STATE"
write-host "====================================="

$filedata = import-csv 'bounce.csv' -Header "date","MID","State"
$filedata | export-csv 'bounce.csv' -NoTypeInformation

write-host "====================================="
write-host "DONE: ADD COLUMN DATE / MID / STATE"
write-host "====================================="

#DELETE UNECESSARY FILE

Remove-Item "balls.csv"
Remove-Item "ball.s"



#CORRELATION OF MESSAGE TRACKING AND BOUNCE
write-host "====================================="
write-host "START: CORRELATION OF STATE (VLOOKUP)"
write-host "====================================="

$data = @{}
Import-Csv 'bounce.csv' | ForEach-Object { $data[$_.MID] = $_.State }
Import-Csv 'all.csv' |  Select-Object *, @{n='State';e={if ($data.Contains($_.MID)) {$data[$_.MID]} else {'success'}}} | export-csv 'gg.csv' -NoTypeInformation
write-host "====================================="
write-host "DONE: CORRELATION OF STATE (VLOOKUP)"
write-host "======================================"
