
# Check and see if already installed
if((Test-Path -Path "C:\Program Files\SpamAssassin\" )){
    $message  = 'Warning'
    $question = "A SpamAssassin installation has been detected already. If you continue all files (including settings) will be overwritten.
    
Are you sure you want to proceed?"

    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

    $decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
    if ($decision -eq 0) {

        # Stop spamd if it exists
        Stop-Service spamd -ErrorAction Ignore

        # Delete spamd if it exists
        $service = Get-WmiObject -Class Win32_Service -Filter "Name='spamd'"
        if($service) {
            $service.delete()
        }

        Remove-Item -Recurse -Force "C:\Program Files\SpamAssassin\"

        schtasks.exe /delete /tn "SpamAssassin AutoUpdate" /F

    } else {
        Exit
    }
}

# Download
Invoke-WebRequest https://downloads.jam-software.de/spamassassin/SpamAssassinForWindows-x64.zip -OutFile C:\Windows\Temp\SpamAssassinForWindows-x64.zip

# Create directory if it doesn't exist
New-Item "C:\Program Files\SpamAssassin\" -type Directory

# Unzip
$shell = new-object -com shell.application
$zip = $shell.NameSpace("C:\Windows\Temp\SpamAssassinForWindows-x64.zip")
foreach($item in $zip.items())
{
    $shell.Namespace("C:\Program Files\SpamAssassin\").copyhere($item)
}

# Downloading srvany-ng

Invoke-WebRequest https://github.com/birkett/srvany-ng/releases/download/v1.0.0.0/srvany-ng_26-03-2015.zip -OutFile C:\Windows\Temp\srvany-ng.zip

# Unzip
Expand-Archive -LiteralPath C:\Windows\Temp\srvany-ng.zip -DestinationPath C:\Windows\Temp\srvany-ng

# Copy to system folder

Copy-Item "C:\Windows\Temp\srvany-ng\x64\srvany-ng.exe" -Destination "C:\Windows\System32\"

#TODO: Get latest version of default configs from the github repo

# Create the service
New-Service -BinaryPathName "C:\Windows\System32\srvany-ng.exe" -Name spamd -DisplayName "SpamAssassin Daemon" 
New-Item -Path HKLM:\SYSTEM\CurrentControlSet\services\spamd -Name "Parameters"
New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\spamd\Parameters -Name "Application" -PropertyType STRING -Value "C:\Program Files\SpamAssassin\spamd.exe"
New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\spamd\Parameters -Name "AppDirectory" -PropertyType STRING -Value "C:\Program Files\SpamAssassin\"
New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\spamd\Parameters -Name "AppParameters" -PropertyType STRING -Value "-x -l -s spamd.log"

# Create the system task to run the update script every night
schtasks.exe /create /tn "SpamAssassin AutoUpdate" /tr "'C:\Program Files\SpamAssassin\sa-update.bat'" /sc DAILY /st 02:00 /RU SYSTEM /RL HIGHEST /v1
schtasks.exe /run /tn "SpamAssassin AutoUpdate"

# wait about 30 seconds for the update to complete
Start-Sleep -Seconds 30

# Download the SpamAssassin config file
Invoke-WebRequest https://raw.githubusercontent.com/winschrott/SpamassassinAgent/master/contrib/spamassassin/local.cf -OutFile "C:\Program Files\SpamAssassin\etc\spamassassin\local.cf"

# Start the spamd service
Start-Service spamd

# Now let's download and install Spamassasssin Transport Agent

# Create Directory to save to
if (!(Test-Path "C:\Customagents")){
  New-Item "C:\CustomAgents\" -type Directory
}

# Create Data Directory
if (!(Test-Path "C:\CustomAgents\SpamassassinAgentData")){
  New-Item "C:\CustomAgents\SpamassassinAgentData" -type Directory
}

# Download the proper DLL
Invoke-WebRequest https://raw.githubusercontent.com/winschrott/SpamassassinAgent/master/bin/SpamassassinAgent.dll -OutFile "C:\CustomAgents\SpamassassinAgent.dll"
Invoke-WebRequest https://raw.githubusercontent.com/winschrott/SpamassassinAgent/master/bin/Microsoft.Exchange.Data.Common.dll -OutFile "C:\CustomAgents\Microsoft.Exchange.Data.Common.dll"
Invoke-WebRequest https://raw.githubusercontent.com/winschrott/SpamassassinAgent/master/bin/Microsoft.Exchange.Data.Common.xml -OutFile "C:\CustomAgents\Microsoft.Exchange.Data.Common.xml"
Invoke-Webrequest https://raw.githubusercontent.com/winschrott/SpamassassinAgent/master/bin/Microsoft.Exchange.Data.Transport.dll -OutFile "C:\CustomAgents\Microsoft.Exchange.Data.Transport.dll"
Invoke-Webrequest https://raw.githubusercontent.com/winschrott/SpamassassinAgent/master/bin/Microsoft.Exchange.Data.Transport.xml -OutFile "C:\CustomAgents\Microsoft.Exchange.Data.Transport.xml"

# Download the XML configuration
Invoke-WebRequest https://raw.githubusercontent.com/winschrott/SpamassassinAgent/master/etc/SpamassassinConfig.xml -OutFile "C:\CustomAgents\SpamassassinAgentData\SpamassassinConfig.xml"

## Connect to the exchange Server
#. 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'
#Connect-ExchangeServer -auto

## Install the Transport Agent
#Install-TransportAgent -Name "SpamAssassin Agent" -AssemblyPath C:\CustomAgents\SpamassassinAgent.dll -TransportAgentFactory SpamassassinAgent.SpamassassinAgentFactory
#Enable-TransportAgent "Spamassassin Agent"
##Set-TransportAgent "Spamassassin Agent" -Priority 3

## Install Anti-Spam Functionality
##. 'C:\Program Files\Microsoft\Exchange Server\V15\Scripts\Install-AntiSpamAgents.ps1'

## Disable Existing Anti-Spam Functionality
##Set-ContentFilterConfig -Enable $false
##Set-SenderFilterConfig -Enable $false
##Set-SenderIDConfig -Enable $false
##Set-SenderReputationConfig -Enable $false

## Restart 
#Restart-Service MSExchangeTransport
