import-module $PSScriptRoot\Dicom.Cmdlet
$conf = Read-Properties $PSScriptRoot\csv2dcm.properties
$outdir = ($PSScriptRoot +"\out\")
get-childitem $outdir -Filter *.dcm | foreach { 
	import-dicom $_.FullName | send-dicom -AET $conf.aetitle -SOPClassProvider $conf.destination
	
	}