Add-AzureRmAccount
Select-AzureRmSubscription -Subscription "[sub-id]"

#Must run elevated
$vms = get-azurermvm
foreach($vm in $vms)
{
#Managed Disk
$osdisk = $null
    try {$osdisk = Get-AzureRmDisk -DiskName $vm.StorageProfile.OsDisk.Name -ResourceGroupName $vm.ResourceGroupName -ErrorAction SilentlyContinue
        $vmCreationTime = $($osdisk.TimeCreated)}
    catch {}
    
    if($osdisk -ne $null)
    { $outstring1 =  "$vmCreationTime,$($vm.Name),$($vm.ResourceGroupName)"  
    Add-Content -Path "c:\out.txt" $outstring1
    }

    if($osdisk -eq $null)
    {
    #Unmanaged Disk
        $vmosdiskloc = $vm.StorageProfile.OsDisk.vhd.Uri
        $vmosdiskstorageact = $vmosdiskloc.Substring(8).Split(".")[0]  #extract out the storage account name by removing the https:// then finding the first part before period (.)
        $storaccount = Get-AzureRmStorageAccount | where {$_.StorageAccountName -eq $vmosdiskstorageact}
        $vmosdiskcontainer = $vmosdiskloc.Substring(8).Split("/")[1]  #extract the middle path between location and vhd which would be the container
        $vmosdiskname = $vmosdiskloc.Substring(8).Split("/")[2]
        $OSBlob = Get-AzureStorageBlob -Context $storaccount.Context -Container $vmosdiskcontainer -Blob $vmosdiskname 
        try {$vmCreationTime = [datetime]::ParseExact(($OSBlob.Name.Substring(($OSBlob.Name.Length-18),14)),'yyyyMMddHHmmss',$null)}
        catch { $vmCreationTime = $null}
        if($vmCreationTime -ne $null)
        {  $outstring2 = "$vmCreationTime,$($vm.Name),$($vm.ResourceGroupName)"
        Add-Content -Path "c:\out.txt" $outstring2
         }
    }

}

