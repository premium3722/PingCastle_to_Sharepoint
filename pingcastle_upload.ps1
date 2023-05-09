echo "####################################"
echo "#Version 5.1                      #"
echo "#Install and execute PingCastle	  #"
echo "#Author: Mike Thurm               #"
echo "#Editor: Levin von Känel          #"
echo "#                                 #"
echo "####################################"


#####################################################################################
#Variablen
    $urlpingcastle = "https://github.com/vletoux/pingcastle/releases/download/3.0.0.3/PingCastle_3.0.0.3.zip"   # Von hier wird das File heruntergeladen
    $path = "C:\Pingcastle\PingCastle.zip"	# Ziel verzeichniss

	  $urlSP = "YOUR_SP_URL"	# Hier wird der Report hochgeladen
	  $SPFolder = "DESTINATION_SP_FOLDER"  #Example: "Freigegebene Dokumente/PingCastle"

    $clientid = "YOUR_CLIENT_ID_FORM_SP"
    $clientsecret = "YOURSECRETKEY"
		
#####################################################################################

# Test if pingcastle is already installed
if (Test-Path "C:\PingCastle")
{
    echo "PingCastle is already installed"
}
else 
{
      try 
      {
        #Save the current value of SecurityProtocol
        $previousSecurityProtocol = [Net.ServicePointManager]::SecurityProtocol
        [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
      }
    catch 
    {
        Write-host -ForegroundColor Red "TLS1.2 konnte nicht aktiviert werden"
    }

    if(Get-PackageProvider -Name Nuget -ErrorAction SilentlyContinue) 
    {
    # NuGet ist installiert
    Write-host -ForegroundColor Green "NuGet is installed."
    }
    else 
    {
        # NuGet ist nicht installiert
        Write-host -ForegroundColor Orange "NuGet is not installed."
        try 
        {
            #Install Powershell module
            #Save the current value of SecurityProtocol
            Install-PackageProvider -Name Nuget -MinimumVersion 2.8.5.201 -Force
            sleep 10
            Install-Module PowerShellGet -MinimumVersion 1.6 -Force -AllowClobber
            sleep 10
            echo "Installiere NUGET"

        }
        catch 
        {
            Write-host -ForegroundColor Red "Error during the installation of NuGet module"
            exit 100
        }
    }
    
    if (Get-Module -Name "SharePointPnPPowerShellOnline" -ListAvailable) 
    {
        Write-Host "SharePointPnPPowerShellOnline module is installed" -ForegroundColor Green
    } 
    else 
    {
        try 
        {
         echo "Installiere SP Modul"
         set-psrepository -name "PSgallery" -installationpolicy trusted
            Install-Module SharepointPnPPowershellonline
         sleep 10
         }
        catch 
        {
           Write-host -ForegroundColor Red "Error during the installation of SP Powershell module"
           exit 200
         }
     }

    #Create Directory
    echo "Erstelle C:\Pingcastle Stuktur"
    New-Item "C:\PingCastle" -itemType Directory
    New-Item "C:\PingCastle\Reports" -ItemType Directory
    New-Item "C:\PingCastle\PingCastle" -ItemType Directory

    #Download PingCastle
    echo "Lädt Pingcastle herunter"
    Start-BitsTransfer -Source $urlpingcastle -Destination $path

    #Wait while downloading file
    sleep 15

    #Extract files
    echo "Zip wird enpackt"
    Expand-Archive -Path "C:\PingCastle\PingCastle.zip" -DestinationPath "C:\PingCastle\PingCastle\"
    Remove-Item -Path "C:\PingCastle\PingCastle.zip"
    
    echo "PingCastle successfully installed"
    #Set the SecurityProtocol setting back to the previous value
    [Net.ServicePointManager]::SecurityProtocol = $previousSecurityProtocol
}

#####################################################################################

#Execute PingCastle & upload to Sharepoint
#test if pingcastle is already installed
if (Test-Path "C:\PingCastle") {

    #change Directory to Pingcastle folder
    try {
        cd "C:\PingCastle\PingCastle\"
    }
    catch {
        Write-Error -Message "the path c:\Pingcastle\Pingcastle could not be found"
        exit
    }

    #Update pingcastle
    echo "Update PingCastle...."
    try {
        .\PingCastleAutoUpdater.exe | echo
    }
    catch {
        Write-Error -Message "Cannot update PingCastle"
    }

    #Start PingCastle & create log file
    echo "Starting PingCastle...."
    try {
        .\PingCastle.exe --healthcheck --level Full --log | echo
    }
    catch {
        Write-Error -Message "Cannot run PingCastle"
        exit
    }
    
    echo "Create PingCastle Report in PingCastle\Report"

    #get file
    $filename = Get-ChildItem -Path C:\PingCastle\PingCastle\*.html -Name

    #check if file already exists
    if (Test-Path -Path "C:\PingCastle\Reports\$filename") {
      echo "Altes Pingcastle File wird gelöscht"
        remove-item -Path "C:\PingCastle\Reports\$filename"
    }

    #move file to Reports folder und neue Variable setzen
    Move-Item "$filename" "C:\PingCastle\Reports\$filename"
    $file_path = "C:\PingCastle\Reports\$filename"
    $env:PnPLEGACYMESSAGE='false' 
 

    #connect to sharepoint
    echo "Connect to Sharepoint...."
    try 
    {
        Connect-PnPOnline -Url $urlSP -ClientId $clientid -ClientSecret $clientsecret
    }
    catch 
    {
        Write-Error -Message "Cannot connect to Sharepoint"
    }

    #Upload file
    echo "Upload file to sharepoint"
    try 
    {
        $Upload = Add-PnPFile -Folder $SPFolder -Path $file_path
    }
    catch 
    {
        Write-Error -Message "Cannot Upload file to Sharepoint"
        exit
    }
	Write-host -ForegroundColor Green "File Uploaded!"
}
else 
{
    Write-host -ForegroundColor Red "There was a error Please Start the Skript again"
}


