import-module $PSScriptRoot\Dicom.Cmdlet
$conf = Read-Properties $PSScriptRoot\csv2dcm.properties

$dirtmp =  ($PSScriptRoot +"\tmp\")
$templatefile = ($PSScriptRoot +"\template.dotx")
$outdir = ($PSScriptRoot +"\out\")
$studies = Import-CSV $conf.csvfile -Delimiter ";"
$json2docx = ($PSScriptRoot +"\bin\json2docx.exe")
$magick = ($PSScriptRoot +"\bin\magick.exe")
if (!(Test-Path $dirtmp))
	{ 
		New-Item -ItemType directory -Path $dirtmp
	}
if (!(Test-Path $outdir)) {
	New-Item -ItemType directory -Path $outdir
}
foreach ($study in $studies) {
 
	$fileroot=$dirtmp + $study.studyInstanceUID

	#if ($study.birthDate) { $study.birthDate =  [datetime]::parseexact($study.birthDate , 'dd/MM/yyyy', $null)  }
	#if ($study.studyDate) { $study.studyDate = [datetime]::parseexact($study.studyDate , 'dd/MM/yyyy', $null)  }
 	#$birthdate = $study.birthDate;
	
	#$studyDate = $study.StudyDate;
	#($birthDate).toString('yyyyMMdd')
	#$studyDate
	#$birthdate="20010101"
	#$studydate="20010101"
	#$gender=""
	#if ($study.sex) {
#		$gender=$study.sex
#	}
	try {
		$birthdate=([datetime]::parseexact($study.birthDate, $conf.formatdate, $null)).toString('yyyyMMdd')
	}catch {
	 
	}
	try {
	$studydate=([datetime]::parseexact($study.studyDate , $conf.formatdate, $null)).toString('yyyyMMdd')
	}catch {
	 
	}
	ConvertTo-Json -InputObject $study | Out-File -filepath ($fileroot + ".json") -Encoding ascii
	#Write-Host ("C:\workspace\json2docx\bin\json2docx -t " + $templatefile  + " -j " + ($fileroot + ".json") +" -d " + ($fileroot + ".docx") + " -e png")
	&$json2docx  -t $templatefile -j ($fileroot + ".json") -d ($fileroot + ".docx") -e png
	&$magick convert -channel RGB ($fileroot + ".png") -negate  ($fileroot + ".neg.png")
	ConvertTO-Dicom ($fileroot + ".neg.png") ($fileroot + ".dcm")
	import-dicom  ($fileroot + ".dcm") | 
	edit-dicom -Tag "0010,0020" -Value $study.patientId |
	edit-dicom -Tag "0010,0010"  -Value ($study.lastName +"^" + $study.firstName)|
	 %{if ($study.sex) { 
		edit-dicom -DicomFile $_ -Tag "0010,0040" -Value $study.sex 
	 	}  			else {$_}   } |
 
 	 %{if ($birthdate) { 
		edit-dicom  -DicomFile $_  -Tag "0010,0030" -Value $birthdate  
	 	}	  else {$_}	} |
	edit-dicom -Tag "0008,0060" -Value $study.modalitiesInStudy |
	edit-dicom -Tag "0008,0061" -Value $study.modalitiesInStudy |
	 %{if (-Not $study.accessionNumber::IsNullOrEmpty) { 
		edit-dicom  -DicomFile $_  -Tag "0008,0050" -Value $study.accessionNumber  
	 	} else {$_}} |
		
     %{if ($studydate) { 
		edit-dicom  -DicomFile $_  -Tag "0008,0020"  -Value $studydate  
	 	} else {$_} 	}  |
	 
	 %{if (-Not $study.accessionNumber::IsNullOrEmpty) { 
		edit-dicom  -DicomFile $_  -Tag "0020,0010" -Value $study.accessionNumber 
	 	}  else {$_} 		} |
		
	edit-dicom -Tag "0020,000D" -Value $study.studyInstanceUID  |
	edit-dicom  -Tag "0020,000E" -Value ($study.studyInstanceUID + ".1.0.1") |
	edit-dicom  -Tag "0008,0018" -Value ($study.studyInstanceUID + ".1.0.2") |
	save-dicom -FileName ($outdir+$study.studyInstanceUID +".dcm")
}

