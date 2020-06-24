function Update-ExoPictureFromAD {
    <#
    .SYNOPSIS
        This CMDLET Updates thumbnailPhoto property on EXO from AD
    .DESCRIPTION
        This CMDLET is used to perform a Push Request From Active Directory and update Photo property of the users on EXO (Exchange Online) using the thumbnailPhoto property on AD
    .EXAMPLE
        --------------------------------------------------------------------------------------
        Update-ExoPictureFromAD 
        Update-ExoPictureFromAD -First 0 -Last 20
        Update-ExoPictureFromAD -FN 0 -LN 20
        Update-ExoPictureFromAD -First 0
        Update-ExoPictureFromAD -Last 20
        --------------------------------------------------------------------------------------
    .INPUTS
        None
    .OUTPUTS
        Logs are created on Desktop file location of server where the script was ran
    .NOTES
        FUTURE MODS LISTED BELOW
        1. Error handling for ranges entered if First number is greater than Last number
    #>

    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(Mandatory = $false,
            HelpMessage = "Please Enter the first number where upload will start from !")]
        [Alias("FN")]
        [int32]
        $First,

        [Parameter(Mandatory = $false,
            HelpMessage = "Please Enter the last number where upload will stop at !")]
        [Alias("LN")]
        [int32]
        $Last
    )

    #log path
    $LogPath = $Home + "\Desktop\Update-ExoPictureFromAD"

    #start log
    Start-Transcript -OutputDirectory $LogPath

    #Get All AD Users
    $Users = Get-ADUser -Filter * -Properties *

    #Set Variable for Total Mailbox in Organisation
    $TotalNumber = $Users.Count

    $i = 0

    if (-Not $First) {
        $FirstValue = 0
    }
    else {
        $FirstValue = $First
    }

    if (-Not $Last) {
        $LastValue = $TotalNumber
    }
    else {
        $LastValue = $Last
    }
    
    #echo number of users in the organisation 
    Write-Host "The Total Number of Users in the Organisation is " $TotalNumber -ForegroundColor yellow

    #echo values of the range of users that upload is been done for
    Write-Host "Starting Upload for Users From the ranges of " $FirstValue "to " $LastValue -ForegroundColor white

    #Region WorkingLoop
    Foreach ($User in $Users) {
        $i++
        if (($i -ge $FirstValue) -and ($i -le $LastValue)) {

            #echo number count of mailbox
            Write-Host "starting for " $i 
            
            #set variable to thumbnailPhoto
            $ThumbPhoto = $User.thumbnailPhoto

            #set variable to UserPrincipalName
            $UserPName = $User.UserPrincipalName

            #set variable to Name
            $UName = $User.Name

            #check block 1 for thumbnailPhoto and UserPrincipalName
            if ((-Not $ThumbPhoto) -and (-Not $UserPName)) {
            
                #echo info 1 for thumbnailPhoto and UserPrincipalName
                Write-Host $UName " has no thumbnailPhoto and UserPrincipalName property" -ForegroundColor Red

                #check block 2 for thumbnailPhoto and UserPrincipalName
            }
            elseif (($ThumbPhoto) -and ($UserPName)) {
            
                #echo info 2 for thumbnailPhoto and UserPrincipalName
                Write-Host $UName " has both thumbnailPhoto and UserPrincipalName property" 

                #check for mailbox
                $Mailbox = Get-Mailbox -Identity $UserPName -ErrorAction SilentlyContinue | Format-List

                #mailbox if block check 1
                if (-Not $Mailbox) {
                
                    #echo mailbox info 1
                    Write-Host "Could not find mailbox for the User " $UserPName -ForegroundColor Red
            
                    #mailbox if block check 2
                }
                elseif ($Mailbox) {
                
                    #echo mailbox info 2
                    Write-Host $UserPName " Mailbox found"

                    #echo update photo info
                    Write-Host "Trying to update thumbnailPhoto for " $UserPName " in Exchange Online" -ForegroundColor yellow

                    #set photo on cloud
                    $SetPhoto = Set-UserPhoto -Identity $UserPName -PictureData ($ThumbPhoto) -Confirm:$false

                    #start sleep cmdlet
                    Start-Sleep -m 0.5;

                    #echo photo info 2
                    Write-Host $UserPName " was updated successfully" -ForegroundColor Green

                }
                #check block 3 for thumbnailPhoto and UserPrincipalName
            }
            elseif ((-Not $ThumbPhoto) -and ($UserPName)) {
            
                #echo info 3 for thumbnailPhoto and UserPrincipalName
                Write-Host $UName " has no thumbnailPhoto property but has UserPrincipalName property" -ForegroundColor Gray
            }
        }
    }
    #EndRegion WorkingLoop

    #stop log
    Stop-Transcript
}