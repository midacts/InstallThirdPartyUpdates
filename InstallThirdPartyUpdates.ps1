#################
### VARIABLES ###
#################
$domain = "# [DOMAINNAME]"
$username = "# [USERNAME]"

# MACHINE PINGLIST
#
# This gets a list of all the online PCs so we can start remote sessions with # powershell
$desk = get-adcomputer -searchbase "# [MACHINE OU]" -filter *
$name = @()
foreach($mach in $desk){$name += ($mach).name}
$name = @($name | ? {@("# LIST OF COMPUTERS TO REMOVE FROM LIST") -notcontains $_})
$pinglist = $name | sort-object
# Here is goes!
$truism = @()
$deskfalicy = @()
$godesk = @()
foreach($comp in $pinglist){
$test = test-connection -count 1 -computer $comp -quiet
$truism += "$comp is $test"
if($test -eq $TRUE){$godesk += $comp}
if($test -eq $FALSE){
write-host "$comp is offline...blast!"
$deskfalicy += $comp
}
}

# Initiates remote powershell sessions
#
# This imports our list of all online computers, and uses your credentials to # start remote powershell sessions on the online PCs.

$comp = $godesk
$cred = Get-credential -credential $domain\$username
$session = New-Pssession -cn $comp -credential $cred -authentication Credssp



# This is the script that will actually import the functions that will be run # on the remote PCs
#
#

Invoke-command -Session $session -Scriptblock {
# Install Third Party Updates Script
# Author: John Patrick McCarthy
# Date: 18th March, 2013
# Version 1.2
#
# To God only wise, be glory through Jesus Christ forever. Amen.
#
# Romans 16:27 ; I Corinthians 15:1-4
#----------------------------------------------------------------
###############
###VARIABLES###
###############
$comp = gc env:computername
$dir = "# [ROOT DIRECTORY WHERE ALL THIRD PARTY MSIs ARE STORED]"
$logdir = "# [DIRECTORY WHERE ALL LOG FILES ARE STORED]"
Function JAVA{
$javaver = '1.7.0_17'
$JavaTest = (test-path "hklm:\software\wow6432node\javasoft\java runtime environment\$javaver")
$JavaTest86 = (test-path "hklm:\software\javasoft\java runtime environment\$javaver")
if($JavaTest -or $JavaTest86 -eq $TRUE){
break
}
if($JavaTest -or $JavaTest86 -eq $FALSE){
$prog = "_Java"
$inst = $("$dir" + "Java\Java.msi")
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="Java.mst" /qn /L $tempfile | out-null
#log name
$javaname = get-itemproperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\4EA42A62D9304AC4784BF238120771FF\InstallProperties"
$info = ($javaname).displayname
$ver = $info -replace "Java ", ""
$logver = $("$logdir" + "$comp" + "$prog" + "($ver)" + ".txt")
out-file $logver
get-content $tempfile | set-content $logver
remove-item $tempfile
}
}
Function FLASH-ACTIVEX{
$activexver = '11_6_602_180'
$path = "C:\Windows\SysWOW64\Macromed\Flash\FlashUtil32_$activexver_ActiveX.exe"
$path86 = "C:\Windows\System32\Macromed\Flash\FlashUtil32_$activexver_ActiveX.exe"
$Flash_Activex = test-path $path
$Flash_Activex86 = test-path $path86
if($Flash_ActiveX -or $Flash_ActiveX86 -eq $TRUE){
break
}
if($Flash_ActiveX -or $Flash_ActiveX86 -eq $FALSE){
$prog = "_Flash ActiveX"
$inst = "$("$dir" + "Flash\Active X\Activex.msi")"
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="Active.mst" /qn /L $tempfile | out-null
#log name
$Flash_Activex = test-path $path
$Flash_Activex86 = test-path $path86
if ($Flash_Activex -eq $TRUE){$activex = gci $path}
if ($Flash_Activex86 -eq $TRUE){$activex = gci $path86}
$info = ($activex).versioninfo.productversion
$ver = $info -replace ",","."
$logver = $("$logdir" + "$comp" + "$prog" + "($ver)" + ".txt")
out-file $logver
get-content $tempfile | set-content $logver
remove-item $tempfile
}
}
Function FLASH-PLUGIN{
$pluginver = '11_6_602_180'
$path = "C:\Windows\SysWOW64\Macromed\Flash\FlashPlayerPlugin_$pluginver.exe"
$path86 = "C:\Windows\System32\Macromed\Flash\FlashPlayerPlugin_$pluginver.exe"
$Flash_Plugin = test-path $path
$Flash_Plugin86 = test-path $path86
if($Flash_Plugin -or $Flash_Plugin86 -eq $TRUE){
break
}
if($Flash_Plugin -or $Flash_Plugin86 -eq $FALSE){
$prog = "_Flash Plugin"
$inst = $("$dir" + "Flash\Plugin\Plugin.msi")
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="Plugin.mst" /qn /L $tempfile | out-null
#log name
$Flash_Plugin = test-path $path
$Flash_Plugin86 = test-path $path86
if ($Flash_Plugin -eq $TRUE){$plugin = gci $path}
if ($Flash_Plugin86 -eq $TRUE){$plugin = gci $path86}
$info = ($plugin).versioninfo.productversion
$ver = $info -replace ",","."
$logver = $("$logdir" + "$comp" + "$prog" + "($ver)" + ".txt")
out-file $logver
get-content $tempfile | set-content $logver
remove-item $tempfile
}
}
Function ADOBE-READER{
$readerver = "11.0.02"
$mspver = "AcroRead11002"
#Change: msiexec.exe /i $inst TRANSFORMS="AcroRead.mst" /update $upd /qn /L $tempfile | out-null
$reader = test-path "c:\program files (x86)\adobe\reader *"
$reader86 = test-path "c:\program files\adobe\reader *"
if($reader -eq $TRUE){($reader_version = (gci "c:\program files (x86)\adobe\reader *").name) -and ($version_Adobe_Reader = (gci "c:\program files (x86)\adobe\$reader_version\reader\acrord32.exe").versioninfo.productversion)}
if($reader86 -eq $TRUE){($reader_version86 = (gci "c:\program files\adobe\reader *").name) -and ($version_Adobe_Reader86 = (gci "c:\program files\adobe\$reader_version86\reader\acrord32.exe").versioninfo.productversion)}
if($version_Adobe_Reader -eq $NULL){$FindReader = $version_Adobe_Reader86}
if($version_Adobe_Reader86 -eq $NULL){$FindReader = $version_Adobe_Reader}
$ReadTest = ($FindReader -lt "$readerver")
if($ReadTest -eq $TRUE){
$prog = "_Adobe Reader"
$inst = $("$dir" + "Adobe Reader\AcroRead.msi")
$upd = $("$dir" + "Adobe Reader\$mspver.msp")
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="AcroRead.mst" /update $upd /qn /L $tempfile | out-null
#log file
$reader = test-path "c:\program files (x86)\adobe\reader *"
$reader86 = test-path "c:\program files\adobe\reader *"
if($reader -eq $TRUE){($reader_version = (gci "c:\program files (x86)\adobe\reader *").name) -and ($version_Adobe_Reader = (gci "c:\program files (x86)\adobe\$reader_version\reader\acrord32.exe").versioninfo.productversion)}
if($reader86 -eq $TRUE){($reader_version86 = (gci "c:\program files\adobe\reader *").name) -and ($version_Adobe_Reader86 = (gci "c:\program files\adobe\$reader_version86\reader\acrord32.exe").versioninfo.productversion)}
if($version_Adobe_Reader -eq $NULL){$FindReader = $version_Adobe_Reader86}
if($version_Adobe_Reader86 -eq $NULL){$FindReader = $version_Adobe_Reader}
$logver = $("$logdir" + "$comp" + "$prog" + "($findreader)" + ".txt")
out-file $logver
get-content $tempfile | set-content $logver
remove-item $tempfile
}
}
Function FIREFOX{
$firefoxver = "17.0.4"
$firefox = test-path "C:\program files (x86)\mozilla firefox"
$firefox86 = test-path "C:\program files\mozilla firefox"
if($firefox -or $firefox86 -eq $TRUE){
if($firefox -eq $TRUE){$Version_firefox = (gci "c:\program files (x86)\mozilla firefox\firefox.exe").versioninfo.productversion}
if($firefox86 -eq $TRUE){$version_Firefox86 = (gci "c:\program files\mozilla firefox\firefox.exe").versioninfo.productversion}
if($Version_firefox -eq $NULL){$FindFirefox = $Version_firefox86}
if($Version_firefox86 -eq $NULL){$FindFirefox = $Version_firefox}
$FirefoxTest = ($FindFirefox -lt $firefoxver)
if($FirefoxTest -eq $FALSE){
break
}
if($FirefoxTest -eq $TRUE){
$prog = "_Firefox"
$inst = $("$dir" + "Firefox\Firefox.msi")
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="Firefox.mst" /qn /L $tempfile | out-null
#log files
$firefox = test-path "C:\program files (x86)\mozilla firefox"
$firefox86 = test-path "C:\program files\mozilla firefox"
if($firefox -eq $TRUE){$Version_firefox = (gci "c:\program files (x86)\mozilla firefox\firefox.exe").versioninfo.productversion}
if($firefox86 -eq $TRUE){$version_Firefox86 = (gci "c:\program files\mozilla firefox\firefox.exe").versioninfo.productversion}
if($Version_firefox -eq $NULL){$FindFirefox = $Version_firefox86}
if($Version_firefox86 -eq $NULL){$FindFirefox = $Version_firefox}
$logver = $("$logdir" + "$comp" + "$prog" + "($findfirefox)" + ".txt")
out-file $logver
get-content $tempfile | set-content $logver
remove-item $tempfile
}
}
}
Function CHROME{
$chromever = "26.0.1410.43"
$Chrome = test-path "C:\Program Files (x86)\Google\Chrome"
$Chrome86 = test-path "C:\Program Files\Google\Chrome"
if($Chrome -or $Chrome86 -eq $TRUE){
if($Chrome -eq $TRUE){$Version_Chrome = (gci "C:\Program Files (x86)\Google\Chrome\application\chrome.exe").versioninfo.productversion}
if($Chrome86 -eq $TRUE){$Version_Chrome86 = (gci "C:\Program Files\Google\Chrome\application\chrome.exe").versioninfo.productversion}
if($Version_Chrome -eq $NULL){$FindChrome = $Version_Chrome86}
if($Version_Chrome86 -eq $NULL){$FindChrome = $Version_Chrome}
$ChromeTest = ($FindChrome -lt $chromever)
if($ChromeTest -eq $FALSE){
break
}
if($ChromeTest -eq $TRUE){
$prog = "_Chrome"
$inst = $("$dir" + "Chrome\GoogleChromeStandaloneEnterprise.msi")
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst /qn /L $tempfile | out-null
$Chrome = test-path "C:\Program Files (x86)\Google\Chrome"
$Chrome86 = test-path "C:\Program Files\Google\Chrome"
if($Chrome -eq $TRUE){$Version_Chrome = (gci "C:\Program Files (x86)\Google\Chrome\application\chrome.exe").versioninfo.productversion}
if($Chrome86 -eq $TRUE){$Version_Chrome86 = (gci "C:\Program Files\Google\Chrome\application\chrome.exe").versioninfo.productversion}
if($Version_Chrome -eq $NULL){$FindChrome = $Version_Chrome86}
if($Version_Chrome86 -eq $NULL){$FindChrome = $Version_Chrome}
$logver = $("$logdir" + "$comp" + "$prog" + "($Findchrome)" + ".txt")
out-file $logver
get-content $tempfile | set-content $logver
remove-item $tempfile
}
}
}
Function SAFARI{
$safariver = "5.1.7 (7534.57.2)"
$Safari = test-path "C:\Program Files (x86)\Safari\"
$Safari86 = test-path "C:\Program Files\Safari\"
if($Safari -or $Safari86 -eq $TRUE){
if($Safari -eq $TRUE){$version_Safari = (gci "c:\program files (x86)\safari\safari.exe").versioninfo.productversion}
if($Safari86 -eq $TRUE){$version_Safari86 = (gci "c:\program files\safari\safari.exe").versioninfo.productversion}
if($version_Safari -eq $NULL){$FindSafari = $version_Safari86}
if($version_Safari86 -eq $NULL){$FindSafari = $version_Safari}
$SafariTest = ($FindSafari -lt $safariver)
if($SafariTest -eq $FALSE){
break
}
if($SafariTest -eq $TRUE){
$prog = "_Safari"
$inst = $("$dir" + "Safari\Safari.msi")
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="Safari.mst" /qn /L $tempfile | out-null
$Safari = test-path "C:\Program Files (x86)\Safari\"
$Safari86 = test-path "C:\Program Files\Safari\"
if($Safari -eq $TRUE){$version_Safari = (gci "c:\program files (x86)\safari\safari.exe").versioninfo.productversion}
if($Safari86 -eq $TRUE){$version_Safari86 = (gci "c:\program files\safari\safari.exe").versioninfo.productversion}
if($version_Safari -eq $NULL){$FindSafari = $version_Safari86}
if($version_Safari86 -eq $NULL){$FindSafari = $version_Safari}
$logver = $("$logdir" + "$comp" + "$prog" + "($findsafari)" + ".txt")
out-file $logver
get-content $tempfile | set-content $logver
remove-item $tempfile
}
}
}
Function ITUNES{
$itunesver = "11.0.2.26"
$Itunes = test-path "C:\Program Files (x86)\iTunes\Itunes.exe"
$Itunes86 = test-path "C:\Program Files\iTunes\itunes.exe"
if($Itunes -or $Itunes86 -eq $TRUE){
if($Itunes -eq $TRUE){$version_Itunes = (gci "c:\program files (x86)\itunes\itunes.exe").versioninfo.productversion}
if($Itunes86 -eq $TRUE){$version_Itunes86 = (gci "c:\program files\itunes\itunes.exe").versioninfo.productversion}
if($version_itunes -eq $NULL){$FindItunes = $version_itunes86}
if($version_itunes86 -eq $NULL){$FindItunes = $version_itunes}
$ItunesTest = ($FindItunes -lt $itunesver)
if($ItunesTest -eq $FALSE){
break
}
if($ItunesTest -eq $TRUE){
#Bonjour
$prog = "_Itunes_Bonjour"
if($Itunes -eq $TRUE){$inst = $("$dir" + "Itunes\64 bit\Bonjour\Bonjour64.msi")}
if($Itunes86 -eq $TRUE){$inst = $("$dir" + "Itunes\32 bit\Bonjour\Bonjour.msi")}
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="Bonjour.mst" /qn /L $tempfile | out-null
$pack = (get-content $tempfile) | select-string "error status: ." | select -expand matches
$val = $pack.value -replace ":",""
$logsp = $("$logdir" + "$comp" + "$prog" + "($val)" + ".txt")
out-file $logsp
get-content $tempfile | set-content $logsp
remove-item $tempfile

#AppSupport
$prog = "_Itunes_Appsupport"
if($Itunes -eq $TRUE){$inst = $("$dir" + "Itunes\64 bit\AppSupport\AppleApplicationSupport.msi")}
if($Itunes86 -eq $TRUE){$inst = $("$dir" + "Itunes\32 bit\AppSupport\AppleApplicationSupport.msi")}
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="Appsupport.mst" /qn /L $tempfile | out-null
$pack = (get-content $tempfile) | select-string "error status: ." | select -expand matches
$val = $pack.value -replace ":",""
$logsp = $("$logdir" + "$comp" + "$prog" + "($val)" + ".txt")
out-file $logsp
get-content $tempfile | set-content $logsp
remove-item $tempfile

#MobileDevice
$prog = "_Itunes_Mobile"
if($Itunes -eq $TRUE){$inst = $("$dir" + "Itunes\64 bit\MobileDevice\AppleMobileDeviceSupport64.msi")}
if($Itunes86 -eq $TRUE){$inst = $("$dir" + "Itunes\32 bit\MobileDevice\AppleMobileDeviceSupport.msi")}
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="MobileDevice.mst" /qn /L $tempfile | out-null
$pack = (get-content $tempfile) | select-string "error status: ." | select -expand matches
$val = $pack.value -replace ":",""
$logsp = $("$logdir" + "$comp" + "$prog" + "($val)" + ".txt")
out-file $logsp
get-content $tempfile | set-content $logsp
remove-item $tempfile

#ITUNES
$prog = "_Itunes_Itunes"
if($Itunes -eq $TRUE){$inst = $("$dir" + "Itunes\64 bit\Itunes\ITunes64.msi")}
if($Itunes86 -eq $TRUE){$inst = $("$dir" + "Itunes\32 bit\Itunes\ITunes.msi")}
$tempfile = $("$logdir" + "$comp" + "$prog" + ".txt")
msiexec.exe /i $inst TRANSFORMS="ITunes.mst" /qn /L $tempfile | out-null
$Itunes = test-path "C:\Program Files (x86)\iTunes\Itunes.exe"
$Itunes86 = test-path "C:\Program Files\iTunes\itunes.exe"
if($Itunes -eq $TRUE){$version_Itunes = (gci "c:\program files (x86)\itunes\itunes.exe").versioninfo.productversion}
if($Itunes86 -eq $TRUE){$version_Itunes86 = (gci "c:\program files\itunes\itunes.exe").versioninfo.productversion}
if($version_itunes -eq $NULL){$FindItunes = $version_itunes86}
if($version_itunes86 -eq $NULL){$FindItunes = $version_itunes}
$logver = $("$logdir" + "$comp" + "$prog" + "($finditunes)" + ".txt")
out-file $logver
get-content $tempfile | set-content $logver
remove-item $tempfile
}
}
}
}

Invoke-command -Session $session -Scriptblock {JAVA}
Invoke-command -Session $session -Scriptblock {FLASH-ACTIVEX}
Invoke-command -Session $session -Scriptblock {FLASH-PLUGIN}
Invoke-command -Session $session -Scriptblock {ADOBE-READER}
Invoke-command -Session $session -Scriptblock {FIREFOX}
Invoke-command -Session $session -Scriptblock {CHROME}
Invoke-command -Session $session -Scriptblock {SAFARI}
Invoke-command -Session $session -Scriptblock {ITUNES}