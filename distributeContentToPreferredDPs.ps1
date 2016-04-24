import-module ConfigurationManager

$site = ""

cd $site

function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	
	$nameWithExtension = "$($OpenFileDialog.filename).csv"
	return $nameWithExtension
}

#open a file dialog window to save the output
$fileName = Save-File $fileName

$storeMe = get-cmapplication | select CI_ID
$i = 0

foreach ($app in $storeMe.CI_ID) {
	Try {
		cd "$site:"
		set-cmapplication -ID $app -sendtoprotecteddistributionpoint 1
		$props = [ordered]@{
			'Application' = $app
			'Status' = 'Ok'
			'Details' = ''			
			}
		$obj = New-Object -TypeName PSObject -Property $props
	}
	
	Catch {
			$ErrorMessage = $_.Exception.Message
			$props = [ordered]@{
			'Application' = $app
			'Status' = 'Error'
			'Details' = $ErrorMessage
			}
			
		$obj = New-Object -TypeName PSObject -Property $props
	}
	
	Finally {
		$data = @()
		$data += $obj
		cd "C:\Powershell"
		$data | Export-Csv $filename -noTypeInformation -append
	}
	$i++
	Write-Progress -activity "Updating application $i of $($storeMe.count)" -percentComplete ($i / $storeMe.Count*100)
}