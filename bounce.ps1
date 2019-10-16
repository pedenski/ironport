
#BOUNCE TO CSV



Cat *.s | sc ball.s #Consolidate all .S
write-host "[1] Concatenate Bounce to CSV format"
(Get-Content 'ball.s') | Where-Object { $_ -match 'careserror' } | Set-Content 'ball.s' #Check per line and get only lines with carerserror
write-host "[2] Remove non-careserror"
Set-Content -Path "ball.s" -Value (get-content -Path "ball.s" | Select-String -Pattern 'delayed' -NotMatch) #Remove lines with delayed 
write-host "[3] Remove Delayed"
(Get-content 'ball.s') | ForEach-Object { $_ -replace ',',':::'} | Set-Content 'ball.s' #replace all comma (,) with (:::) this avoids CSV to delimit comma and separate it per column.
write-host "[4] Foreach , replace :::"
import-csv ball.s | export-csv balls.csv  -NoTypeInformation #convert *.s by exporting to csv


#GETTING MID AND STATE FROM BOUNCE
write-host "[5] Instantiate Excel"
$thispath = Get-Location
#write-host "opening $thispath\balls.csv"
$excel = new-object -comobject Excel.Application
$excel.visible = $false
$workbook = $excel.workbooks.open("$thispath\balls.csv")
$worksheet = $workbook.Worksheets.Item(1)

write-host "[6] Count Rows"
$colCount = $worksheet.range("A1").currentregion.rows.count #get Column count of column A
$worksheet.range("B1:B$colCount").formula = '=MID(A1,FIND("MID",A1)+4,9)' #add formula to columb B
write-host "[7] Apply getMID"
$worksheet.range("C1:C$colCount").formula = 
'=IF(ISNUMBER(SEARCH("spam content blocked",A1)),"rejected by destination", IF(OR(ISNUMBER(SEARCH({"does not exist","is not exist","doesnt have","aborted","expired","no such user","mailbox unavailable"},A1))),"invalid email",IF(OR(ISNUMBER(SEARCH({"dns soft","dns hard"},A1))),"invalid domain",IF(ISNUMBER(SEARCH("disabled",A1)),"mailbox disabled",IF(ISNUMBER(SEARCH("quota",A1)),"mailbox full","rejected by destination")))))'
write-host "[8] Apply getState"
$workbook.SaveAs("$thispath\bounce.csv")
$workbook.Close($false)
$excel.Quit()



# need to release COM references and ensures no remaining references exist which will keep Excel.exe open 
# Without performing the above, if you look in Task Manager, you may see Excel still running...in some cases, many copies.
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Remove-Variable -Name excel
write-host "[9] Release comObjects"

#ADD COLUMN DATE MID STATE


write-host "[10] Add Column / Mid / State"
$filedata = import-csv 'bounce.csv' -Header "date","MID","State"
$filedata | export-csv 'bounce.csv' -NoTypeInformation

#DELETE UNECESSARY FILE


Remove-Item "balls.csv"
Remove-Item "ball.s"

write-host "[11] Cleanup balls"


#CORRELATION OF MESSAGE TRACKING AND BOUNCE
write-host "[12] Execute correlation vlookup"

$data = @{}
Import-Csv 'bounce.csv' | ForEach-Object { $data[$_.MID] = $_.State }
Import-Csv 'sorted.csv' |  Select-Object *, @{n='State';e={if ($data.Contains($_.MID)) {$data[$_.MID]} else {'success'}}} | export-csv 'gg.csv' -NoTypeInformation

write-host "[13] Bounce cleanup"
Remove-Item "bounce.csv"
Remove-Item "sorted.csv"
write-host "[14] GG!!"