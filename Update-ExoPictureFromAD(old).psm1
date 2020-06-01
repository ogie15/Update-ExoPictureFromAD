function Update-ExoPictureFromAD () {
    #log path
    $LogPath= $Home+"\Desktop\Update-ExoPictureFromAD.txt"

    #start log
    Start-Transcript -Path $LogPath

    #Get All AD Users
    $Users = Get-ADUser -Filter * -Properties *

    Foreach ($User in $Users) {
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
        }elseif (($ThumbPhoto) -and ($UserPName)) {
        
            #echo info 2 for thumbnailPhoto and UserPrincipalName
            Write-Host $UName " has both thumbnailPhoto and UserPrincipalName property" 

            #check for mailbox
            $Mailbox = Get-Mailbox -Identity $UserPName -ErrorAction SilentlyContinue | Format-List

            #mailbox if block check 1
            if (-Not $Mailbox) {
            
                #echo mailbox info 1
                Write-Host "Could not find mailbox for the User " $UserPName -ForegroundColor Red
        
            #mailbox if block check 2
            }elseif ($Mailbox) {
            
                #echo mailbox info 2
                Write-Host $UserPName " Mailbox found"

                #echo update photo info
                Write-Host "Trying to update thumbnailPhoto for " $UserPName " in Exchange Online" -ForegroundColor yellow

                #set photo on cloud
                $SetPhoto = Set-UserPhoto -Identity $UserPName -PictureData ($ThumbPhoto) -Confirm:$false

		        Start-Sleep -m 0.5;

                #echo photo info 2
                Write-Host $UserPName " was updated successfully" -ForegroundColor Green

            }
        #check block 3 for thumbnailPhoto and UserPrincipalName
        }elseif ((-Not $ThumbPhoto) -and ($UserPName)) {
        
            #echo info 3 for thumbnailPhoto and UserPrincipalName
            Write-Host $UName " has no thumbnailPhoto property but has UserPrincipalName property" -ForegroundColor Gray
        }
    }
    Stop-Transcript
}