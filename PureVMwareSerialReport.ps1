# VMware/Pure Analyzer
# Written by Joshua Woleben


# Pure arrays
$pure_arrays = @("purearray_1","pure_array2")
$pure_volumes= @{}
$datastore_to_pure = @{}
# Vcenter host
$vhosts = @("vcenter_host1","vcenter_host2")

$excel_file = "C:\Temp\VMwarePureSerialNumber_$(get-date -f MMddyyyyHHmmss).xlsx"
$TranscriptFile = "C:\Temp\VMwarePureSerialNumber_$(get-date -f MMddyyyyHHmmss).txt"
Start-Transcript -Path $TranscriptFile
Write-Output "Initializing..."

# Import required modules
Import-Module PureStoragePowerShellSDK
Import-Module VMware.VimAutomation.Core

# Define a gigabyte in bytes
$gb = 1073741824

# Gather authentication credentials
Write-Output "Please enter the following credentials: `n`n"

# Collect vSphere credentials
Write-Output "`n`nvSphere credentials:`n"
$vsphere_user = Read-Host -Prompt "Enter the user for the vCenter host"
$vsphere_pwd = Read-Host -Prompt "Enter the password for connecting to vSphere: " -AsSecureString
$vsphere_creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $vsphere_user,$vsphere_pwd -ErrorAction Stop

$pure_user = Read-Host -Prompt "Enter the user for the Pure storage arrays"
$pure_pwd = Read-Host -Prompt "Enter the password for the Pure storage array user: " -AsSecureString
$pure_creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $pure_user,$pure_pwd -ErrorAction Stop

# Open Excel


$erroractionpreference = "SilentlyContinue"
$excel_object = New-Object -comobject Excel.Application
$excel_object.visible = $True 

$excel_workbook = $excel_object.Workbooks.Add()
$excel_worksheet = $excel_workbook.Worksheets.Item(1)

# Create headers
$excel_worksheet.Cells.Item(1,1) = "VMware Datastore Name"
$excel_worksheet.Cells.Item(1,2) = "VMware Datastore UUID"
$excel_worksheet.Cells.Item(1,3) = "Pure Volume Name"
$excel_worksheet.Cells.Item(1,4) = "Pure Volume UUID"
$excel_worksheet.Cells.Item(1,5) = "Pure Array Name"
$excel_worksheet.Cells.Item(1,6) = "Inventory match?"
$d = $excel_worksheet.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$row_counter = 2

# Get all pure volumes on all arrays
Write-Host "Gathering Pure volumes..."
ForEach ($array in $pure_arrays) {

    # Connect to Pure Array
    $pure_connect = New-PfaArray -EndPoint $array -Credentials $pure_creds -IgnoreCertificateError -ErrorAction Stop

    # Get all volumes
    $pure_volumes[$array] += Get-PfaVolumes -Array $pure_connect

    # Disconnect Pure array
    Disconnect-PfaArray -Array $pure_connect

}

foreach ($vcenter_host in $vhosts) {
    # Connect to vCenter
    Connect-VIServer -Server $vcenter_host -Credential $vsphere_creds -ErrorAction Stop

    # Get All VMs
    # Write-Host "Gathering all VMs..."
    # $vm_collection = Get-VM -Server $vcenter_host

    # Get all datastores
    Write-Host "Gathering datastores..."
    $datastore_collection = Get-Datastore -Server $vcenter_host





    # Figure out what array a datastore is on
    Write-Host "Determining datastore array locations..."
    $datastore_collection | ForEach-Object {

        # Get disk name
        $disk_name = $_.Name

        # Get UUID from VMware
        $uuid = $_.ExtensionData.Info.Vmfs.Extent[0].DiskName

        Write-Host "Processing $disk_name..."

        # Translate VMware UUID to Pure UUID by removing the naa. and the first eight characters, and converting to uppercase
        $pure_uuid = ($uuid -replace "naa\.\w{8}","").ToUpper()
        Write-Host "UUID: $uuid Pure UUID: $pure_uuid"


        # Search each array for the Pure UUID
        ForEach ($array in $pure_arrays) {

            # Search each volume for the correct UUID
            $pure_volumes[$array] | ForEach-Object { 
                # If UUID found, store with array name
                if (($_ | Select -ExpandProperty serial) -eq $pure_uuid) {
                            Write-Host "$disk_name found on $array!"
                            $datastore_to_pure[$disk_name] = $pure_uuid

                            $excel_worksheet.Cells.Item($row_counter,1) = $disk_name
                            $excel_worksheet.Cells.Item($row_counter,2) = $uuid
                            $pure_volume_name = ($_ | Select -ExpandProperty name)
                            $pure_volume_uuid = ($_ | Select -ExpandProperty serial)

                            $excel_worksheet.Cells.Item($row_counter,3) = $pure_volume_name
                            $excel_worksheet.Cells.Item($row_counter,4) = $pure_volume_uuid

                            $excel_worksheet.Cells.Item($row_counter,5) = $array

                            if (("$disk_name" -notmatch "$pure_volume_name") -and ("$pure_volume_name" -notmatch "$disk_name")) {
                                $excel_worksheet.Cells.Item($row_counter,6) = "No"
                                $row_range = $excel_worksheet.Range("A$row_counter"+":F$row_counter")
                                $row_range.Interior.ColorIndex = 3
                            }
                            else {
                                $excel_worksheet.Cells.Item($row_counter,6) = "Yes"                                
                                $row_range = $excel_worksheet.Range("A$row_counter"+":F$row_counter")
                                $row_range.Interior.ColorIndex = 4

                            }

                            $row_counter++
                }

            }
        }
    
    }
    # Disconnect from vCenter
    Disconnect-VIServer -Server $vcenter_host -Confirm:$false
}

#Format Excel

$e = $excel_worksheet.Range("A1:F$row_counter")
$e.Borders.Item(12).Weight = 2
$e.Borders.Item(12).LineStyle = 1
$e.Borders.Item(12).ColorIndex = 1

$e.Borders.Item(11).Weight = 2
$e.Borders.Item(11).LineStyle = 1
$e.Borders.Item(11).ColorIndex = 1

$e.BorderAround(1,4,1)

$e.Columns("A:F").AutoFit()

$sort_range = $excel_worksheet.Range("F1")
$e.Sort($sort_range,1)


# Save Excel
$excel_workbook.SaveAs($excel_file)
$excel_workbook.Close
$excel_object.Quit()

[System.RunTime.IntropServices.Marshall]::ReleaseComObject($excel_object)
Remove-Variable $excel_object

# Generate email report
$email_list=@("user1@example.com","user2@example.com")
$subject = "Pure/VMware Serial Report"

$body = "Serial report attached."

Stop-Transcript

$MailMessage = @{
    To = $email_list
    From = "SerialReport<Donotreply@example.com>"
    Subject = $subject
    Body = $body
    SmtpServer = "smtp.example.com"
    ErrorAction = "Stop"
    Attachment = $excel_file
}
Send-MailMessage @MailMessage

