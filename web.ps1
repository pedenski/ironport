$username = "a-zdmurai"
$password = "P@ssword01" 

$ie = New-Object -com InternetExplorer.Application 
$ie.visible=$true
$ie.navigate("https://10.31.20.80/") 
while($ie.ReadyState -ne 4) {start-sleep -m 100};
if ($ie.document.url -Match "invalidcert")
        {
        "Bypassing SSL Certificate Error Page";
        $sslbypass=$ie.Document.IHTMLDocument3_getElementsByTagName("a") | where-object {$_.id -eq "overridelink"};
        $sslbypass.click();
        "sleep for 3 seconds while final page loads";
        start-sleep -s 3;
        };
if ($ie.Document.domain -Match "10.31.20.80")
        {
        "Successfully Bypassed SSL Error";
        }
        else
        {
        "Bypass failed";
        }


$form =  $ie.document.forms[0]
$inputs = $form.getElementsByTagName("input")

($inputs | where {$_.name -eq "username"}).value = $username
($inputs | where {$_.name -eq "password"}).value = $password
($inputs | where {$_.name -eq "action:Login"}).click()

while($ie.ReadyState -ne 4 -or $ie.Busy) {Start-Sleep -m 1000}
Start-Sleep -m 2000


$ie.navigate("https://10.31.20.80/monitor/message_tracking")
write-host "accessing mesage tracking"
while($ie.ReadyState -ne 4 -or $ie.Busy) {Start-Sleep -m 2000}


$form =  $ie.document.forms[0]
$form = ($ie.document.forms | where {$_.id -eq "trackingSearchForm"})

$inputs = $form.getElementsByTagName("input")
$select = $form.getElementsByTagName("select")





($inputs | where {$_.name -eq "sender"}).value = "careserror@pldt.com.ph"
($select | where {$_.name -eq "subject_match"}).value = "match_contains"



($inputs | where {$_.name -eq "subject"}).value = "aug 26"
($inputs | where {$_.id -eq "submitButton"}).click()


while($ie.ReadyState -ne 4) {start-sleep -m 900};



#$postParams = @{action='Seach';sender_match='match_begins';sender='careserror@pldt.com.ph';subject_match='match_contains';subject='aug 26'}
#Invoke-WebRequest -Uri https://10.31.20.80/ -Method POST -Body $postParams


if ($ie.document.url -Match "message_tracking")
        {
        while($ie.ReadyState -ne 4) {start-sleep -m 900};
        write-host "1"
       

       <# if ($az=$ie.Document.IHTMLDocument3_getElementsByTagName("a") |? {$_.textContent -Match "Export"}) {
        $az.click()
       
        write-host "1" #>

        } else {
        write-host "0"
        }
       
       
       




<#$ie.document.getElementByID("sender").value = "test"
$ie.document.getElementByID("match_contains").value = "test2"#>

<#
$form =  $ie.document.forms[0]
$inputs = $form.getElementsByTagName("input")
($inputs | where {$_.name -eq "sender"}).value = "test"#>