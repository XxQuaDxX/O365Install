##Run Complete Office Cleaner###########################################################################################################
Start-Process -FilePath cscript -ArgumentList "offscrub03.vbs STANDARD,Personal.PROPLUS,XLView,WDView,VisView,HtmlView /L /quiet /nocancel /bypass 1" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process -FilePath cscript -ArgumentList "offscrub07.vbs STANDARD,Personal.PROPLUS,XLView,WDView,VisView,HtmlView,BASIC,HomeAndStudent,Ultimate,PPTView,Personal,SmallBusiness /L /quiet /nocancel /bypass 1" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process -FilePath cscript -ArgumentList "offscrub10.vbs STANDARD,Personal.PROPLUS,XLView,WDView,VisView,HtmlView,BASIC,HomeAndStudent,Ultimate,STARTER,PptView,CLICK2RUN,SmallBusiness,Personal,ACADEMIC /L /quiet /nocancel /bypass 1" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process -FilePath cscript -ArgumentList "offscrub13ctr.vbs all /L /quiet /nocancel /bypass 1" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process -FilePath cscript -ArgumentList "offscrubc2r.vbs all /L /quiet /nocancel /bypass 1" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process -FilePath cscript -ArgumentList "offscrub_o15msi.vbs STANDARD,ProPlus,XLView,WDView,VisView,HtmlView,BASIC,HomeAndStudent,Ultimate,STARTER,PptView,CLICK2RUN,SmallBusiness,Personal,ACADEMIC /L /quiet /nocancel /bypass 1" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process -FilePath cscript -ArgumentList "offscrub_o16msi.vbs STANDARD,ProPlus,XLView,WDView,VisView,HtmlView,BASIC,HomeAndStudent,Ultimate,STARTER,PptView,CLICK2RUN,SmallBusiness,Personal,ACADEMIC /L /quiet /nocancel /bypass 1" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue


##Run Office 365 Installation###########################################################################################################
Start-Process .\setup.exe -ArgumentList "/configure configuration.xml" -Wait -WindowStyle Hidden


##Modify Office 365 Install###########################################################################################################
if(Test-Path 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive for Business.lnk'){
         Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive for Business.lnk" -Force -Recurse -ErrorAction SilentlyContinue
}
If(Test-Path 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook 2016.lnk'){
         Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook 2016.lnk" -Force -Recurse -ErrorAction SilentlyContinue
}
If(Test-Path 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\skype for Business 2016.lnk'){
         Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\skype for Business 2016.lnk" -Force -Recurse -ErrorAction SilentlyContinue
}
If(Test-Path 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools\Skype for Business Recording Manager.lnk'){
         Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools\Skype for Business Recording Manager.lnk" -Force -Recurse -ErrorAction SilentlyContinue
}


##Remove Microsoft OneDrive App###########################################################################################################
if(Test-Path 'C:\Program Files (x86)\Microsoft Office\root\Integration\OneDriveSetup.exe'){
         Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\root\Integration\OneDriveSetup.exe" -ArgumentList "/Uninstall" -Wait -ErrorAction SilentlyContinue
}elseif(Test-Path 'C:\Program Files\Microsoft Office\root\Integration\OneDriveSetup.exe'){
         Start-Process -FilePath "C:\Program Files\Microsoft Office\root\Integration\OneDriveSetup.exe" -ArgumentList "/Uninstall" -Wait -ErrorAction SilentlyContinue
}
If(Test-Path 'C:\Program Files\Microsoft OneDrive\OneDriveSetup.exe'){
         Start-Process -FilePath "C:\Program Files\Microsoft OneDrive\OneDriveSetup.exe" -ArgumentList "/Uninstall" -Wait -ErrorAction SilentlyContinue
         Remove-Item -Path "C:\Program Files\Microsoft OneDrive" -Force -Recurse -ErrorAction SilentlyContinue

}elseif(Test-Path 'C:\Program Files (x86)\Microsoft OneDrive\OneDriveSetup.exe'){
         Start-Process -FilePath "C:\Program Files (x86)\Microsoft OneDrive\OneDriveSetup.exe" -ArgumentList "/Uninstall" -Wait -ErrorAction SilentlyContinue
         Remove-Item -Path "C:\Program Files (x86)\Microsoft OneDrive" -Force -Recurse -ErrorAction SilentlyContinue
}
Start-Process .\ProfileClean3.bat -ArgumentList "C:\Users Links\OneDrive for Business.lnk" -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process .\ProfileClean1.bat -ArgumentList "C:\Users Links\OneDrive.lnk" -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process .\ProfileClean1.bat -ArgumentList "C:\Users OneDrive" -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process .\ProfileClean3.bat -ArgumentList "C:\Users OneDrive for Business" -WindowStyle Hidden -ErrorAction SilentlyContinue
If(Test-Path 'C:\Users\Default\Links\OneDrive.lnk'){
         Remove-Item -Path "C:\Users\Default\Links\OneDrive.lnk" -Force -Recurse -ErrorAction SilentlyContinue
}


##Install OWA Shortcuts###########################################################################################################
$owapath1 = "C:\"
$owaICO = $owapath1 + "OUTLOOK2016.ico"
$testowaico = Test-Path $owaICO
If($testowaico -eq $false){
        Copy-Item "OUTLOOK2016.ico" $owapath1 -Force -ErrorAction SilentlyContinue ##Copies .ico file to the root of C:\
    }
$owapath2 = "C:\Users\Public\Desktop"
$owapath3 = "C:\Users\Public\Desktop\Radiology Shortcuts"
$owaICON = $owapath2 + "\Outlook E-Mail.lnk"
$owaradICON = $owapath3 + "\Outlook E-Mail.lnk"
$testowaicon = Test-Path $owaICON
$radowaicon = Test-Path $owaradICON
$chrome32 = Test-Path 'C:\Program Files\Google\Chrome\Application\chrome.exe'
$chrome64 = Test-Path 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
If($testowaicon -eq $false -and $chrome32 -eq $true){
        Copy-Item "OWA32.lnk" $owapath2 -ErrorAction SilentlyContinue ##Copies OWA Shortcut to public desktop
        Rename-Item "C:\Users\Public\Desktop\OWA32.lnk" "Outlook E-Mail.lnk"
    }
If($testowaicon -eq $false -and $chrome64 -eq $true){
        Copy-Item "OWA64.lnk" $owapath2 -ErrorAction SilentlyContinue ##Copies OWA Shortcut to public desktop
        Rename-Item "C:\Users\Public\Desktop\OWA64.lnk" "Outlook E-Mail.lnk"
    }
If($radcpu -eq $true -and $radowaicon -eq $false){
        Copy-Item "OWA64.lnk" $owapath3 -ErrorAction SilentlyContinue ##Copies OWA Shortcut to Radiology Shortcuts
        Rename-Item "C:\Users\Public\Desktop\Radiology Shortcuts\OWA64.lnk" "Outlook E-Mail.lnk"
    }


##Copy New Custome Script###########################################################################################################
$radcpu = Test-Path 'C:\Users\Public\Desktop\Radiology Shortcuts'
If($radcpu -eq $false){
        Remove-Item -Path C:\Masters\DriveShortcuts\CustomScript.vbs -Force -ErrorAction SilentlyContinue
        Copy-Item -Container ".\DriveShortcuts" "C:\Masters" -Force -ErrorAction SilentlyContinue
        Copy-Item ".\DriveShortcuts\*.*" "C:\Masters\DriveShortcuts" -Force -ErrorAction SilentlyContinue
        Start-Process .\ProfileCopy.bat -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
}
If($radcpu -eq $true){
Start-Process .\ProfileClean4.bat -ArgumentList "C:\Users Desktop\Radiology Shortcuts\Microsoft Excel 2010.lnk" -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process .\ProfileClean4.bat -ArgumentList "C:\Users Desktop\Radiology Shortcuts\Microsoft PowerPoint 2010.lnk" -WindowStyle Hidden -ErrorAction SilentlyContinue
Start-Process .\ProfileClean4.bat -ArgumentList "C:\Users Desktop\Radiology Shortcuts\Microsoft Word 2010.lnk" -WindowStyle Hidden -ErrorAction SilentlyContinue
}


##Remove Groupwise Shortcuts###########################################################################################################
$gw = Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Novell GroupWise\GroupWise.lnk"
If($gw -eq $true){
        Remove-Item "C:\Users\Public\Desktop\Novell GroupWise.lnk" -Force -ErrorAction SilentlyContinue
        Remove-Item "C:\Users\Public\Desktop\GroupWise.lnk" -Force -ErrorAction SilentlyContinue
}
If($radcpu -eq $true -and $gw -eq $true){
        Remove-Item "C:\Users\Public\Desktop\Radiology Shortcuts\GroupWise.lnk" -Force -ErrorAction SilentlyContinue
        Remove-Item "C:\Users\Public\Desktop\Radiology Shortcuts\Novell GroupWise.lnk" -Force -ErrorAction SilentlyContinue
}


##Replace Removed Shortcuts###########################################################################################################
If($radcpu -eq $true){ ##Rad Computers places icons in Public Desktop Radiology Shortcuts folder
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint 2016.lnk" "C:\Users\PublicDesktop\Radiology Shortcuts" -Force -ErrorAction SilentlyContinue
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word 2016.lnk" "C:\Users\Public\Desktop\Radiology Shortcuts" -Force -ErrorAction SilentlyContinue
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel 2016.lnk" "C:\Users\Public\Desktop\Radiology Shortcuts" -Force -ErrorAction SilentlyContinue
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher 2016.lnk" "C:\Users\Public\Desktop\Radiology Shortcuts" -Force -ErrorAction SilentlyContinue
}
$OfficeProg = Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office"
If($OfficeProg -eq $false){ ##Creates Microsoft Office Folder and places shortcuts inside (To make it more familiar for End Users)
        mkdir -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office" ##Creates Microsoft Office folder if not there
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint 2016.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office" -Force -ErrorAction SilentlyContinue
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word 2016.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office" -Force -ErrorAction SilentlyContinue
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel 2016.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office" -Force -ErrorAction SilentlyContinue
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher 2016.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office" -Force -ErrorAction SilentlyContinue
        Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote 2016.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office" -Force -ErrorAction SilentlyContinue
}