Add-Type -AssemblyName System.Windows.Forms
Add-type -AssemblyName System.Drawing


function row-model($sTag, $sysname){

$ie = new-object -ComObject "InternetExplorer.Application"
    $proCount = 0
 $ie.Navigate("www.dell.com/support/home/us/en/04/product-support/servicetag/$sTag/configuration")
    while($ie.busy){Start-Sleep -Milliseconds 10}    
    if($ie.document.body.innerHTML -like "*small form facto*") {$curmod = "Optiplex $sysname SFF"}#small form factor
    elseif($ie.document.body.innerHTML -like "*Desktop Base*") {$curmod = "Optiplex $sysname"}#desktop
    elseif($ie.document.body.innerHTML -like "*micro form facto*") {$curmod = "Optiplex $sysname Micro"}#micro
    elseif($ie.document.body.innerHTML -like "*minitower*") {$curmod = "Optiplex $sysname MT"} #Mini Tower
    elseif($ie.document.body.innerHTML -like "*ultra small form facto"){$curmod = "Optiplex $sysname USFF"} #ultra small form factor
    elseif($ie.document.body.innerHTML -like "*aio xcto*") {$curmod = "AIO XCTO"}
    elseif($ie.document.body.innerHTML -like "*base,all in one*") {$curmod = "AIO BTX"}
    elseif($ie.document.body.innerHTML -like "*Support for Latitude*") {$curmod = "Latitude $sysname"}
    else{$curmod = "serial not found"}    
       $ie.Quit()  
     return $curmod
}

function draw-size($parm1,$parm2){
$drawsize = New-Object System.Drawing.Size($parm1, $parm2)
return $drawsize
}

function click-selectfile(){
$SaveChooser = New-Object -TypeName System.Windows.Forms.OpenFileDialog
$SaveChooser.ShowDialog()
$global:openFile = $SaveChooser.FileName
}

function csv-complete($PathToCsv){
$rawCsv = Import-Csv -Delimiter "," -Header @("a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z") -Path $PathToCsv | select a,b,d,e,f,o,i,z |select -skip 1 


$homecsv = "Service Tag, Mac Address, Mode, Stats, Asset Tage, Duplicate, Manufacturer, Purchase Order, Device Type"
$homecsv | out-file -Encoding ascii "reporttest.csv" -Force 

$streamW = New-Object -TypeName System.IO.StreamWriter -ArgumentList ".\reporttest.csv"
#write the headers to the csv.
$streamW.WriteLine( '$assetnames' + ',' + 'MacAddresses' + ',' + 'models'+ ',' + 'Status'+ ',' + 'assetTags'+ ',' + 'serviceTags'+ ',' + 'Manufacturer'+ ',' +  'purchaseOrders'+ ',' +  'deviceTypes')
$repnum = 0
foreach($row in $rawCsv){
$assetNames = $row.B
$macAddresses = $row.Z
$models = $row.O
$Status = "in-service"
$assetTags = $row.A
$serviceTags = $row.B
$Manufacturer = "Dell"
$purchaseOrders = $row.D
$deviceTypes = $row.F
$rowmodel = row-model $row.b $row.O
write-host $assetTags 
write-host $rowmodel
$repnum += 1

$curLine = $assetnames+ ',' + $MacAddresses+ ',' + $rowmodel + ',' + $Status+ ',' + $assetTags+ ',' + $serviceTags+ ',' + $Manufacturer+ ',' +  $purchaseOrders+ ',' +  $deviceTypes
$streamW.WriteLine("$curline")
}
$streamW.close()
}

function draw-form(){
$form = New-Object System.Windows.Forms.Form
$form.Text = "Report Gen"
$selectFileButton = New-Object System.Windows.Forms.Button
$selectFileButton.Text = "Browse"
$selectFileButton.Anchor = "top,left"
$selectFileButton.Add_Click({click-selectfile 
$selectedFileBox.Text = $openFile})

$selectedFileBox = new-object System.Windows.Forms.TextBox
$selectedFileBox.Text = "C:\users\user\folder\report.csv"
$selectedFileBox.Size = draw-size 250 50
$selectedFileBox.Anchor = "top"

$exportcsvbtn = new-object System.Windows.Forms.button
$exportcsvbtn.Text = "Finish"
$exportcsvbtn.Anchor = "top,right"
$exportcsvbtn.Add_Click({csv-complete $selectedFileBox.Text})

$form.controls.Add($selectFileButton)
$form.controls.add($selectedFileBox)
$form.controls.add($exportcsvbtn)

$fileheight = $selectFileButton.Height + 40
$form.Size = draw-size 801 $fileheight
$form.StartPosition = 'centerscreen'
$form.ShowDialog()

return $form
}

draw-form
