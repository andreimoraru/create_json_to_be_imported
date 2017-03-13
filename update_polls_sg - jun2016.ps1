#################################################################################
#    for updating monitoring system items
#    
#  to do:
#    - add multi-threading [IN PROGRESS...]
#    - compile GUI VS solution
#    - code cleanup & obfuscation
#    - rewrite in C#
#
#
#                            Andrei Moraru
#                            Jun 2016
#
#################################################################################


$lgf = "<log file>"
$plf = "<poll file>"
$agmacc = "<creds>"
$agmpw = "<creds>"
$swcmpurist = "http://<replace-here>/welcome.php?switchCompanyID="
$lguri = "<uri>"
$lguribd = "tz=Europe%2FAthens&username=$agmacc&password=$agmpw"
$lgurirsp = Invoke-WebRequest -Uri $lguri -Body $lguribd -Method Post -SessionVariable $ssnvr 
$cmpid = 3
$swcmpuri  = $swcmpurist + $cmpid
$swcmpurirsp   = Invoke-WebRequest -Uri $swcmpuri -WebSession $ssnvr -Method Get
$edpluri = "http://<replace-here>/editPoll.php"
$cn = Get-Content $plf
$ftbl = @()
foreach ($ln in $cn) {
    
    $arroel = ($ln -Split "||", 0, "simplematch")[0,1,2]

    $ftbl += (,($arroel[0],$arroel[1],$arroel[2]))
}

function Release-Ref ($rf) {

([System.Runtime.InteropServices.Marshal]::ReleaseComObject(

[System.__ComObject]$rf) -gt 0)

[System.GC]::Collect()

[System.GC]::WaitForPendingFinalizers()

}
$snm = "Capacity Polls"
$str = 2
$wtt = “<company name>_Polls1.xlsx”
$efl = “xlsx folder path>”
$objExcel = New-Object -ComObject Excel.Application 
$objExcel.Visible = $false
$uwb = $objExcel.Workbooks.Open($efl + $wtt)
$ush = $uwb.Sheets.Item($snm)
Do {

 if ($ush.Cells.Item($str,1).Value()) {

    $pfn = $ush.Cells.Item($str, 3).Value()
    $htn = $ush.Cells.Item($str, 1).Value()
    $pdc = $ush.Cells.Item($str, 7).Value()
    $pln = (($ftbl -match "$htn") -match (("$pfn" -replace "\(","\(") -replace "\)","\)"))
    if($pln.Count -eq 0) {Write-Host "`nError: The spreadsheet poll $pfn for asset $htn wasn't matched against any poll in Augmenta web interface`n" | Tee-Object -Append $lgf} 
            elseif ($pln.Count.Equals(1)) {
                $pid = $pln[0][0]
                $gpldturib = "pollID=$pid"
                $gpldtrs = Invoke-WebRequest -Uri $edpluri -WebSession $ssnvr -Body $gpldturib -Method Post
                if ($gpldtrs.StatusCode -eq 200) {
                    $plto  = $gpldtrs.InputFields.FindByName("timeout").value
                    $plrt = $gpldtrs.InputFields.FindByName("retries").value
                    if($pid -and $pfn -and $plto  -and $plrt ) {
                        $edplurib= "pollID=$pid&timeout=$plto &retries=$plrts&fieldName=$pfn&description=$pdc"
                        $edplurirs = Invoke-WebRequest -Uri $edpluri -WebSession $ssnvr   -Method Post -Body $editPollURIbody
                        $plup = ($edplurirs.AllElements | Where-Object {$_.ID -eq "rightcontent"}).innerText
                        if ($plup -eq "Poll updated") {
                            Write-Output "`nPoll $pfn of $htn was updated to new value for Poll description. This is the new value`n`n`t $pdc`n`n" | Tee-Object -Append $lgf} 
                            else {Write-Output " Error: Poll $pfn of $htn wasn't updated, please review the script source code`n`n" | Tee-Object -Append $lgf}
                        Remove-Variable pid,pfn,plrt,pdc,plto
                    } else {Write-Host "Error: Poll $pfn of $htn wasn't updated, because one of the required poll fields is NULL or missing: Poll ID, Poll Field Name, Poll Timeout, Poll Retries, Poll Description`n" | Tee-Object -Append $lgf}
                    } else {Write-Host "`nError: HTTP Status code for web response if other than 200 OK. Check web response from Augmenta web interface`n"  | Tee-Object -Append $lgf}
                }
                else {Write-Host "`nError: The Augmenta poll $pfn for asset $($pln[0][2]) is matched $($pln.Count) times.`nPlease fix that in Augmenta web interface!`nIt's on row $str in spreadsheet`n"  | Tee-Object -Append $lgf}
    }
           $str++
           } While ($str -le ($ush.UsedRange).SpecialCells([Microsoft.Office.Interop.Excel.Constants]::xlLastCell).Row)
$objExcel.Quit()
$a = Release-Ref($ush)
$a = Release-Ref($uwb) 
$a = Release-Ref($objExcel)
