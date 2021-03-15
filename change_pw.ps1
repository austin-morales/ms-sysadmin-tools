function Change-Password
{
    #Change Password
    $Name = Read-Host "Which user password would you like to change?"
    $Password = Read-Host "What would you like to set their password to?"
    Set-ADAccountPassword -Identity $Name -NewPassword (ConvertTo-SecureString -AsPlainText -String $Password -Force)   
    Write-Host "Done!"
    Write-Host "Sending email..."
}

function Create-Email
{
    #Create/Send Email
    $Outlook = New-Object -ComObject Outlook.Application

    #logic to determine if we're emailing a store or a person
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = "$name@domain.com"
}

function Send-Email
{
    #compose message
    $Mail.Subject = "New Password"
    $Mail.Body ="Hello, $Name's new password is $Password

    If you want to change it, head over to [AD PASSWORD CHANGE SERVICE] and you'll be able to change your password using the one in this email."
    $Mail.Send()
}

function Update-User
{
    #write to user
    Write-Host "Done!"
    Write-Host "$Name 's password is $Password."
    Write-Host ""
    Read-Host "Press enter to quit"
}

Change-Password
Create-Email
Send-Email
Update-User