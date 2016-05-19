Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

Function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	$nameWithExtension = "$($OpenFileDialog.filename).csv"
	return $nameWithExtension
}

Function Get-Folder($initialDirectory) {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.RootFolder = [System.Environment+SpecialFolder]'MyComputer'
    $browse.ShowNewFolderButton = $false
    $browse.Description = "Choose a directory"

    $loop = $true
    while($loop)
    {
        if ($browse.ShowDialog() -eq "OK")
        {
            $loop = $false
        } else
        {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Try again or exit script?", "Choose a directory", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if($res -eq "Cancel")
            {
                return
            }
        }
    }
    $browse.SelectedPath
    $browse.Dispose()
}


Function checkPatches($SystemName){

	# Set up some variables to hold referenced results from Render
	$deviceInfo = "<DeviceInfo><NoHeader>True</NoHeader></DeviceInfo>"
	$extension = ""
	$mimeType = ""
	$encoding = ""
	$warnings = $null
	$streamIDs = $null
	
	$reportPath = "/Software Update Compliance/$ssrsReportName"
	$Report = $RS.GetType().GetMethod("LoadReport").Invoke($RS, @($reportPath, $null))
	
	$parameters = @()
	$parameters += New-Object RS.ParameterValue
	$parameters[0].Name  = "SystemName"
	$parameters[0].Value = "$SystemName"

	$RS.SetExecutionParameters($parameters, "en-us") > $null
	$RenderOutput = $RS.Render('csv',$deviceInfo,[ref] $extension,[ref] $mimeType,[ref] $encoding,[ref] $warnings,[ref] $streamIDs)

	$Stream = New-Object System.IO.FileStream("$folderLoc\$SystemName.csv"), Create, Write
	$Stream.Write($RenderOutput, 0, $RenderOutput.Length)
	$Stream.Close()
}
#SSRS Server
$server = ""
$reportServerURI = "http://$server/ReportServer/ReportExecution2005.asmx?WSDL"
$RS = New-WebServiceProxy -Class 'RS' -NameSpace 'RS' -Uri $reportServerURI -UseDefaultCredential
$RS.Url = $reportServerURI

$serverList = Get-Content -Path (Get-FileName)
#$fileName = Save-File $fileName
$folderLoc = Get-Folder

$i = 0

$erroractionpreference = "SilentlyContinue"

$DEV = ""
$PROD = ""

foreach($server in $serverList){
	$errorMessage = ""
	$i++
	Write-Progress -id 1 -activity "Generating report for: $server `($i of $($serverList.count)`)" -percentComplete ($i / $serverList.Count*100)

	Try{
		$siteCode = (Invoke-WMIMethod -computername $Server -namespace root\ccm -Class SMS_Client -Name GetAssignedSite).sSiteCode
		if($siteCode -eq "$DEV"){
			$ssrsReportName = "Single_Server_Dev"	
			checkPatches($server)
		}
		if($siteCode -eq "$PROD"){
			$ssrsReportName = "Single_Server"
			checkPatches($server)
		}
	}

	Catch{
		$errorMessage = $_.Exception.Message	
	}
	
	$serverInfo = New-Object -TypeName PSObject -Property @{
		Server = $server
		SiteCode = $siteCode
		Details = $errorMessage
	}
	
	$serverInfo | Export-Csv $folderLoc\Logs.csv -noTypeInformation -append
}

