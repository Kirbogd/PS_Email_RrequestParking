<#
PowerShell Script to create a draft for parking request
Asks for car make, Plate number, date (default is set in the beginning of the script, tomorrow and adjusted to Russian plate standards)
and creates email draft to preset address

author: kirbogd@microsoft.com
Version: 1.0
#>

# Let's Define default Make, Plate and Date

$DefMake = "Put_Your_Model_and_Make".ToString()
$DefPlate = "PutYourPlateNumber"
$DefDate = (Get-Date).AddDays(1).ToString('dd.MM.yyy')

#Define email address and Email heading. Message body is set inside a following function

$Recipient = "put_your@Email.here"
$Header = "Put your email header here"

function send-ParkRequest ([string]$CarMake = $DefMake,[string]$Plate = $DefPlate,[string]$Date = $DefDate ) {
    
    ## check parameters format

    if ($CarMake.Length -gt "20") {
        Write-Host "Your Model Name is tooooooooo long" -ForegroundColor Red
        break
    }
    if ($Plate -notmatch "[a-zA-Z]\d{3}[a-zA-Z]{2}\d{2,3}"){
        Write-Host "It is not a ligit Plate number" -ForegroundColor Red
        break 
    }
    if ($Date -notmatch "\d{2}.\d{2}.\d{4}") {
        Write-Host "Please enter date in dd.MM.yyyy format" -ForegroundColor Red
        break 
    }
    
		# Set email text
	
		$MailBodyText = "Hi! <br> May I have a parking lot booked for <br>"+$CarMake+" <br>plate number "+$Plate+" <br>during "+ $Date + "? <br><br>"
    $ol = New-Object -comObject Outlook.Application 
    
    # Get email signature (for Outlook)
    $MailSignature = Get-Content ($env:USERPROFILE + "\AppData\Roaming\Microsoft\Signatures\*.htm")
    
    # call the save method to have the email in the drafts folder
    $mail = $ol.CreateItem(0)
    $null = $Mail.Recipients.Add($Recipient)  
    $Mail.Subject = $Header  
    $Mail.HTMLBody = $($MailBodyText)+$MailSignature
    $Mail.save()
    $inspector = $mail.GetInspector
    $inspector.Display()
    
    Write-host "All is done, check your drafts in Outlook" -ForegroundColor Green
   
}

Write-Host "Set your car model and make (no more 20 characters). Press enter for default -"$DefMake": " -NoNewline -ForegroundColor Green
$UserCar = Read-Host 
Write-Host "Set you car's plate number in A123BC456 format (English layout). Press enter for default -"$DefPlate": " -NoNewline -ForegroundColor Green
$UserPlate = Read-Host 
Write-Host "Set the requested date in dd.MM.yyyy format. Press enter for tomorrow -"$DefDate": " -NoNewline -ForegroundColor Green
$UserDate = Read-Host 

if ($UserCar -eq $null -or $UserCar.Length -lt "1"){
    $UserCar = $DefMake
}
if ($UserPlate -eq $null -or $UserPlate.Length -lt "1"){
    $UserPlate = $DefPlate
}
if ($UserDate -eq $null -or $UserDate.Length -lt "1"){
    ($UserDate = $DefDate)
}

send-ParkRequest -CarMake $UserCar -Plate $UserPlate -Date $UserDate 
