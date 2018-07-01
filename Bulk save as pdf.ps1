# Add the PowerPoint assemblies that we'll need
Add-type -AssemblyName office -ErrorAction SilentlyContinue
Add-Type -AssemblyName microsoft.office.interop.powerpoint -ErrorAction SilentlyContinue

# Start PowerPoint
$ppt = new-object -com powerpoint.application
$ppt.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Set the locations where to find the PowerPoint files, and where to store the thumbnails
$pptPath = "C:\mstemp\Workshop\WsP-NET Fwrk Dev Mod WApps w ASP.NET MVC\PPTs\"


# Loop through each PowerPoint File
Foreach($iFile in $(ls $pptPath -Filter "*.ppt")){
Set-ItemProperty ($pptPath + $iFile) -name IsReadOnly -value $false
$filename = Split-Path $iFile -leaf
$file = $filename.Split(".")[0]
$oFile = $pptPath + $file + ".pdf" 

# Open the PowerPoint file
$pres = $ppt.Presentations.Open($pptPath + $iFile)

# Now save it away as PDF 
$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF 
$pres.SaveAs($ofile,$opt)


# and Tidy-up 
$pres.Close();

}

#Clean Up
$ppt.quit();
$ppt = $null
[gc]::Collect();
[gc]::WaitForPendingFinalizers();