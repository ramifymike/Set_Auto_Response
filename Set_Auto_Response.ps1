##### OUT OF OFFICE PS1 #####
 
function Set-OutOfOffice {
 
  param ( 
  [Parameter(Mandatory=$true, Position=0)] [System.String]${Identity},
  [Parameter(Mandatory=$true, Position=1)] [System.String]${Message}
  )
 
  #Implicit remote to office365
  $getcred = Get-Credential -Message "PLEASE ENTER YOUR OFFICE 365 EMAIL ACCOUNT INFORMATION!!"
 
 
  #Office 365
  $o365 = New-PSSession -ConfigurationName Microsoft.Exchange `
  -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $getcred `
  -Authentication Basic -AllowRedirection
 
  #Import Session
  Import-PSSession $o365 | Out-Null
  #Done
 
  #Set-MailboxReplyConfiguration
  Set-MailboxAutoReplyConfiguration -Identity $Identity -AutoReplyState Enabled `
  -ExternalAudience All -InternalMessage $Message -ExternalMessage $Message
 
}
 
try {
  Set-OutOfOffice $args[0] $args[1]
}
 
catch {
  Write-Output "
  -- Set Out of Office Script --
  Two arguments required: email, message
  Usage Example:
  .\out_of_office.ps1 %email% %message%
  .\out_of_office.ps1 mike@ramify.io "I am out of the office"
  "
}
