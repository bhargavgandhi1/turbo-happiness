. C:\Users\Bhargav\Desktop\python_scripts\Get-NetworkStatistics.ps1

$appName = "CDP"
$filter = "name like '%"+$appName+"%'"
$result = Get-WmiObject win32_Service -Filter $filter 
#Construct an out-array to use for data export
$OutArray = @()
foreach($service in $result)
    {
        $prid = $service.ProcessId
	$instance = $service.Name
	$ip=get-WmiObject Win32_NetworkAdapterConfiguration|Where {$_.Ipaddress.length -gt 1} 
	$hostip = $ip.ipaddress[0]
	$stat = Get-NetworkStatistics | Where-Object {$_.PID -eq $prid} 
	foreach($netout in $stat) {
		$servicelocalip = $stat.LocalAddress
		if ($servicelocalip -eq $hostip) {
			$port = $netout.LocalPort
			#Construct an object
        		$myobj = "" | Select "ServiceName","FixPort","HostName"
        		#fill the object
        		$myobj.ServiceName = $instance
        		$myobj.FixPort = $port
        		$myobj.HostName = $env:computername

		        #Add the object to the out-array
        		$outarray += $myobj

		        #Wipe the object just to be sure
        		$myobj = $null
		}
	}

    }
#After the loop, export the array to CSV
Write-Host $outarray
$outarray | export-csv "somefile.csv"