<#
#    Move-DistributionGroups.ps1
#
#    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
#    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,        
#    INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
#    We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute
#    the object code form of the Sample Code, provided that You agree: (i) to not use Our name, Logo, or trademarks
#    to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on
#    Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us
#    and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the 
#    use or distribution of the Sample Code.
#    Please note: None of the conditions outlined in the disclaimer above will supersede the terms and conditions contained 
#    within the Premier Customer Services Description.
#

.SYNOPSIS
Export all groups/members to CSV files e after creates them in the EXO, finally adds each member in their respective groups
Move-DistributionGroups.ps1

.DESCRIPTION
The Move-DistributionGroups script exports the attributes and members of all distribution groups from the OnPrem to the EXO, moreover adds the members in their respective groups

Obs: You must execute this script with a credential that has acess to all groups onprem and also you should has a credential member of the role needed to create and edit EXO distribution groups

.ATTRIBUTES
The following attributes will be copied

    Name,Alias,BypassNestedModerationEnabled,DisplayName,ManagedBy,MemberDepartRestriction,MemberJoinRestriction,ModeratedBy,ModerationEnabled,SendModerationNotifications,AcceptMessagesOnlyFromDLMembers,AcceptMessagesOnlyFrom,HiddenFromAddressListsEnabled,PrimarySmtpAddress,RejectMessagesFrom,RejectMessagesFromDLMembers,RequireSenderAuthenticationEnabled,EmailAddresses,bypassModerationFromSendersOrMembers,GrantSendOnBehalfTo,SendAsPermission

.PreREQUISITES
You must follow the exection order, including the need to move to the EXO any mailboxes in the fields ManagedBy and ModeratedBy before run the option 3

.PARAMETER Available Options
1 - Export a list containing all the distribution groups and their respective members

The first option exports a list containing all distribution groups to the file DistributionGroups.csv. Also, exports a list of members per group to the file DistributionGroups_Members.csv. This option must be the first one to be executed in the case
of a cut over migration. Any files with the same name in the script directory will be overwritten.

2 - Remove all groups from the Exchange OnPrem

The second option removes all groups listed in the DistributionGroups.csv from the Exchange OnPrem. It is mandatory to run the first option before the second. Is necessary to have the file DistributionGroups.csv in the same path as the script path.

3 - Create all groups in the Exchange Online and add the respective members

The third option creates all groups listed in the file DistributionGroups.csv into EXO, edits them with the attributes previously exported to the file DistributionGroups.csv

.PARAMETER Group

The Group parameter could be used to move a specific group

.LOGGING
The execution of this script will create a Log file in the same directory of this script with the name dd_MM_yyyy.Log, creating a new Log for each new day that the script is executed. This Log will register actions and errors.
The file Log contains the following columns separated by comma:

date = Date in the format dd/MM/yyyy-HH:mm:ss
status = Status of the task, with the following possible values:

ACTION - A action has been executed by the user. One of the available options was selected.
CONN - A connection/session with EXO was created or Remove-Itemeted
ERROR - An erro was found
EXPORT_GROUP - A group was recorded in the file DistributionGroups.csv.
EXPORT_MEMBER - A group member was recorded in the file DistributionGroups_Members.csv.
GROUP_REMOVED - A group was removed from the Exchange OnPrem.
GROUP_CREATED - A group was created in the EXO.
GROUP_EDIT - A group was edited in the EXO.
ADD_MEMBER - A group member was added in a group in the EXO.
message = This column has a detailed message of the executed task

#>

#Parameters
Param(
    [string]$Group
)

#Function to create a Write-Log file and register Write-Log entries
Function Write-Log{
    Param(
        [string]$Status,
        [string]$Message
    )
    
    $logName = Get-Date -Format d_M_yyyy.log

    $dayLogFile = Test-Path $logName
    
    $dateTime = Get-Date -Format dd/MM/yyyy-HH:mm:ss

    If($dayLogFile -eq $true){

        $logLine = $dateTime + "," + $Status + "," + $Message
        $logLine | Out-File -FilePath $logName -Append
    }
    Else
    {
        $header = "Date,Status,Message"
        $header | Out-File -FilePath $logName
        $logLine = $dateTime + "," + $Status + "," + $Message
        $logLine | Out-File -FilePath $logName -Append
    }
}

#Function to connect to the EXO
Function ConnectToEXO{
    
    try{
        $a = Get-Credential
        $a.Password | ConvertFrom-SecureString | Set-Content exo-password.sec

        Write-Log -Status "CONN" -Message "User credentials for the user $a.UserName encripted"

        $userName = $a.UserName
        $password = Get-Content .\exo-contoso.sec | convertto-securestring
        $EXOCredential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password

        $Global:EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $EXOCredential -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue

        Write-Log -Status "CONN" -Message "Session with EXO using the user $userName created sucessfully"
    }
    catch{
        Write-Log -Status "ERROR" -Message "Error when trying to create a EXO session"
        }

    try{
        Import-PSSession $EXOSession -Prefix EXO -CommandName New-DistributionGroup,Add-DistributionGroupMember,Set-DistributionGroup,Get-DistributionGroup,Get-Mailbox,Add-RecipientPermission -AllowClobber -ErrorAction SilentlyContinue
        Write-Log -Status "CONN" -Message "Session with EXO imported sucessfully for the user $userName"
        }
        catch{
            Write-Log -Status "ERROR" -Message "Error when trying to import the EXO session"
        }
}

#Function to export groups
Function Export-Groups{
    Param(
        [string]$GroupList,
        [string]$Group
    )

    Write-Log -Status "ACTION" -Message "Export Groups and Members cmdlet started"

    If($GroupList){
        try{
             $distributionGroups = Import-Csv $GroupList -ErrorAction Stop
            Write-Log -Status "ACTION" -Message "The option to migrate a list of groups from the file $GroupList was selected"
        }
        catch{
            Write-Log -Status "ERROR" -Message "Error when trying to import a list of groups"
        }
    }
    Else
    {
        If($Group){
            try{
                 $distributionGroups = Get-DistributionGroup $Group -ErrorAction Stop
                Write-Log -Status "ACTION" -Message "The option to migrate only the group $Group was selected"
            }
            catch{
                Write-Log -Status "ERROR" -Message "Error when trying to execute the cmdlet Get-DistributionGroup for the group $Group"
            }
        }
        Else
        {
            try{
             $distributionGroups = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop | Select-Object Name,Alias,BypassNestedModerationEnabled,DisplayName,ManagedBy,MemberDepartRestriction,MemberJoinRestriction,ModeratedBy,ModerationEnabled,SendModerationNotifications,AcceptMessagesOnlyFromDLMembers,AcceptMessagesOnlyFrom,HiddenFromAddressListsEnabled,PrimarySmtpAddress,RejectMessagesFrom,RejectMessagesFromDLMembers,RequireSenderAuthenticationEnabled,EmailAddresses,bypassModerationFromSendersOrMembers,GrantSendOnBehalfTo,SendAsPermission
            Write-Log -Status "ACTION" -Message "The option to migrate all distribution groups was selected"
            }
            catch{
            Write-Log -Status "ERROR" -Message "Error when trying to executed the cmdlet Get-DistributionGroup for the group $Group"
            }
        }
    }

    $outputDL = @()
    $outputDLMembers = @()
    $arr = @()

    Foreach($obj in  $distributionGroups){

        $AcceptMessagesOnlyFromDLMembers = $Null
        $AcceptMessagesOnlyFrom = $Null
        $RejectMessagesFrom = $Null
        $RejectMessagesFromDLMembers = $Null
        $EmailAddresses = $Null        
        $managedBy = $Null
        $moderatedBy = $Null
        $bypassModerationFromSendersOrMembers = $Null


        If($obj.ManagedBy){
            $arr = $obj.ManagedBy | Select-Object Name
            foreach($user In $arr){
                $managedBy += $user.Name.ToString() + ";"        
            }
        }
        
        $arr=@()

        If($obj.ModeratedBy){
            $arr = $obj.ModeratedBy | Select-Object Name
            foreach($user In $arr){
                $moderatedBy += $user.Name.ToString() + ";"        
            }
        }        

        $arr=@()

        If($obj.AcceptMessagesOnlyFromDLMembers){
            $arr = $obj.AcceptMessagesOnlyFromDLMembers | Select-Object Name
            foreach($user In $arr){
                $AcceptMessagesOnlyFromDLMembers += $user.Name.ToString() + ";"        
            }
        }

        $arr=@()

        If($obj.AcceptMessagesOnlyFrom){
            $arr = $obj.AcceptMessagesOnlyFrom | Select-Object Name
            foreach($user In $arr){
                $AcceptMessagesOnlyFrom += $user.Name.ToString() + ";"        
            }
        }

        $arr=@()

        If($obj.RejectMessagesFrom){
                $arr = $obj.RejectMessagesFrom | Select-Object Name
                foreach($user In $arr){
                    $RejectMessagesFrom += $user.Name.ToString() + ";"        
                }
            }
        
        $arr=@()

       If($obj.RejectMessagesFromDLMembers){
            $arr = $obj.RejectMessagesFromDLMembers | Select-Object Name
            foreach($user In $arr){
                $RejectMessagesFromDLMembers += $user.Name.ToString() + ";"        
            }
        }
    
        $arr=@()

        If($obj.EmailAddresses){
            $arr = $obj.EmailAddresses | Select-Object SmtpAddress
            foreach($item In $arr){
                If($item -ne ''){
                    $EmailAddresses += $item.SmtpAddress + ";"
                }
            }
        }

        $arr=@()

       If($obj.bypassModerationFromSendersOrMembers){
            $arr = $obj.bypassModerationFromSendersOrMembers | Select-Object Name
            foreach($user In $arr){
                $bypassModerationFromSendersOrMembers += $user.Name.ToString() + ";"        
            }
        }

        $arr=@()
        $GrantSendOnBehalfTo = $Null

        If($obj.GrantSendOnBehalfTo){
            $arr = $obj.GrantSendOnBehalfTo | Select-Object Name
            foreach($user In $arr){
                $GrantSendOnBehalfTo += $user.Name.ToString() + ";"        
            }
        }

        $arr=@()
        $SendAsPermission=””

        $arr = Get-DistributionGroup $obj.Name | Get-ADPermission | Where-Object{($_.ExtendedRights -like “*Send-As*”)}
        
        If($null -eq $arr){
            foreach($user In $arr){
                    If($null -ne $user){
                        $a = $user.User.ToString()
                        $b = $a.substring(0,2)
                        If($b -ne 'S-'){
                            $c = $a.Split("\")
                            $SendAsPermission += $c[1] + ";"
                    }
                }
            }
        }

        $objDL = [PSCustomObject]@{
            Name = $obj.Name
            Alias = $obj.Alias
            BypassNestedModerationEnabled = $obj.BypassNestedModerationEnabled
            DisplayName = $obj.DisplayName
            ManagedBy = $managedBy
            MemberDepartRestriction = $obj.MemberDepartRestriction
            MemberJoinRestriction = $obj.MemberJoinRestriction
            ModeratedBy = $moderatedBy
            ModerationEnabled = $obj.ModerationEnabled
            SendModerationNotifications = $obj.SendModerationNotifications
            MaxSendSize = $obj.MaxSendSize 
            MaxReceiveSize = $obj.MaxReceiveSize
            AcceptMessagesOnlyFromDLMembers = $AcceptMessagesOnlyFromDLMembers
            AcceptMessagesOnlyFrom = $AcceptMessagesOnlyFrom
            HiddenFromAddressListsEnabled = $obj.HiddenFromAddressListsEnabled
            PrimarySmtpAddress = $obj.PrimarySmtpAddress
            RejectMessagesFrom = $RejectMessagesFrom
            RejectMessagesFromDLMembers = $RejectMessagesFromDLMembers
            RequireSenderAuthenticationEnabled =$obj.RequireSenderAuthenticationEnabled
            EmailAddresses = $EmailAddresses
            bypassModerationFromSendersOrMembers = "bypassModerationFromSendersOrMembers"
            GrantSendOnBehalfTo = "GrantSendOnBehalfTo"
            SendAsPermission = "SendAsPermission"
        }

        $outputDL += $objDL
            
        Write-Log -Status "EXPORT_GROUP" -Message "The group $($objDL.Name) was added to export list"

    }

    try{
        $outputDL | Export-Csv DistributionGroups.csv -NoTypeInformation -ErrorAction SilentlyContinue -Encoding UTF8
        Write-Log -Status "EXPORT_GROUP" -Message "The file DistributionGroups.csv was created sucessfully"
    }
    catch{
        Write-Log -Status "ERROR" -Message "Error when trying to create the file GruposDeDistribucao.csv"
    }
        
    Foreach($dg in  $distributionGroups){

        $distributionGroupMembers = Get-DistributionGroupMember $dg.name -resultsize unlimited -ErrorAction SilentlyContinue
        
        Foreach($member in $distributionGroupMembers){
            $groupMember = $member.Alias
            $objMember = [PSCustomObject]@{
                Alias = $member.Alias
                DistributionGroup = $dg.Name
            }
            $outputDLMembers += $objMember
            Write-Log -Status "EXPORT_MEMBER" -Message "The member $groupMember of the group $groupName was added to the export list"
            }
    }

    try{
        $outputDLMembers | Export-Csv DistributionGroups_Members.csv -NoTypeInformation -ErrorAction SilentlyContinue -Encoding UTF8
        Write-Log -Status "EXPORT_GROUP" -Message "The file DistributionGroups_Members.csv was created with sucess"
    }
    catch{
        Write-Log -Status "ERROR" -Message "Error when trying to create the fileDistributionGroups_Members.csv"
    }
}

#Function used to remove groups from the Exchange OnPrem
Function RemoveGroups{

    Param(
        [string]$Group
    )
    
    Write-Log -Status "ACTION" -Message "Option 2 = Remove groups"

    $mode2 = Read-Host "
    WARNING!!! This procedure will remove all groups listed in the file DistributionGroups.csv from the On-Premises organization. Would you like to proceed (y/n)?"
    
    switch($mode2){
        y{
            Write-Log -Status "ACTION" -Message "The exclusion of all groups in the file DistributionGroups.csv was confirmed by the user"

            If($Group){
                try{
                    $dls = Get-DistributionGroup $Group -ErrorAction SilentlyContinue
                    Write-Log -Status "ACTION" -Message "The option to migrate only the group $Group was selected"
                }
                catch{
                    Write-Log -Status "ERROR" -Message $varErro
                    $varErro = ''
                }
            }
            Else{
                $dls = Test-Path DistributionGroups.csv
                $dlsMembers = Test-Path DistributionGroups_Members.csv

                If($dls -eq $true -and $dlsMembers -eq $true){
                    try{
                        $dls = Import-Csv DistributionGroups.csv -ErrorAction SilentlyContinue
                        Write-Log -Status "ACTION" -Message "The option to migrate all groups was selected"
                    }
                    catch{
                        Write-Log -Status "ERROR" -Message $varErro
                        $varErro = ''
                    }
                }
                Else
                {
                    Write-Host "The file DistributionGroups.csv was removed. Keep this file in the same directory as the script or use the option 1 again to recreate them"
                    
                    Write-Log -Status "ERROR" -Message "File DistributionGroups.csv not found"
                }
            }
        
            $dls | ForEach-Object{
                $groupName = $_.Name
                
                try{
                    Remove-DistributionGroup $_.Name -BypassSecurityGroupManagerCheck -Confirm:$y -ErrorAction SilentlyContinue
                    Write-Log -Status "GROUP_REMOVED" -Message "Group $groupName was removed from the Exchange OnPrem"
                }
                catch{
                    Write-Log -Status "ERROR" -Message "Error when trying to remove the group $groupName"
                }
            }
        }
        n{
            Write-Log -Status "ACTION" -Message "Option 2 - Remove Groups aborted"
            Write-Host "Taks aborted!" -ForegroundColor Yellow
        }
    }
}

#Function to create groups in the EXO
Function CreateGroups{

    Param(
        [string]$Group
        )

    Write-Log -Status "ACTION" -Message "Option 3 - Create groups in the EXO was selected"

    $mode2 = Read-Host "
    WARNING!!! Before use the option 3, you must guarantee that the following steps have sucessfully completed:

        1. Only use option 3 before option 2;
        2. Move all users that are either in the field ManagedBy or ModeratedBy to the EXO;
        3. Run the sync, confirm that all groups have been removed from the Windows Azure Active Directory and make sure no error is Write-Logged.

    Proceed ONLY if all 3 steps above have sucessfully completed

    Would you like to proceed (y/n)?"

    switch($mode2){
        y{
            Write-Log -Status "ACTION" -Message "The groups creation in the EXO was selected"

            $dls = Test-Path DistributionGroups.csv
            $dlsMembers = Test-Path DistributionGroups_Members.csv

            If($dls -eq $true -and $dlsMembers -eq $true){
                try{
                $dls = Import-Csv DistributionGroups.csv -ErrorAction SilentlyContinue
                Write-Log -Status "ACTION" -Message "The option to migrate all groups was selected"
            }
            catch{
                Write-Log -Status "ERROR" -Message "Error when trying to import the file DistributionGroups.csv"
                }
            }
            Else
            {
                Write-Host "The file DistributionGroups.csv was removed. Keep the file in the same directory as the script directory or use the option 2 to create them."
                    
                Write-Log -Status "ERROR" -Message "Arquivo DistributionGroups.csv não encontrado"
                
                Exit
            }

            ConnectToEXO

            foreach($dl in $dls){

                switch($dl.BypassNestedModerationEnabled){
                    TRUE{$BypassNestedModerationEnabled = $true}
                    FALSE{$BypassNestedModerationEnabled = $false}
                }

                switch($dl.RequireSenderAuthenticationEnabled){
                    TRUE{$RequireSenderAuthenticationEnabled = $true}
                    FALSE{$RequireSenderAuthenticationEnabled = $false}
                }

                switch($dl.HiddenFromAddressListsEnabled){
                    TRUE{$HiddenFromAddressListsEnabled = $true}
                    FALSE{$HiddenFromAddressListsEnabled = $false}
                }
                
                $groupName = $dl.Name

                try{
                    New-EXODistributionGroup -Name $dl.Name -Alias $dl.Alias -BypassNestedModerationEnabled $BypassNestedModerationEnabled -DisplayName $dl.DisplayName -MemberDepartRestriction $dl.MemberDepartRestriction -MemberJoinRestriction $dl.MemberJoinRestriction -SendModerationNotifications $dl.SendModerationNotifications -Type Distribution | Out-Null
                    Write-Log -Status "GROUP_CREATED" -Message "Group $groupName in the EXO was created sucessfully"
                }
                catch{
                    Write-Log -Status "ERROR" -Message "Error when trying to create the group $groupName in the EXO"
                }
                
                try{
                    Set-EXODistributionGroup $dl.Name -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -PrimarySmtpAddress $dl.PrimarySmtpAddress -ErrorAction SilentlyContinue | Out-Null
                    Write-Log -Status "GROUP_EDIT" -Message "The attributes RequireSenderAuthenticationEnabled, HiddenFromAddressListsEnabled and PrimarySmtpAddress were changedto $RequireSenderAuthenticationEnabled, $HiddenFromAddressListsEnabled and $dl.PrimarySmtpAddress respectively in the group $groupName sucessfully in the EXO."
                }
                catch{
                    Write-Log -Status "ERROR" -Message "Error when trying to change the attributes RequireSenderAuthenticationEnabled, HiddenFromAddressListsEnabled and PrimarySmtpAddress in the grup $groupName in the EXO."
                }

                If($dl.ManagedBy){
                    $arr = $dl.ManagedBy.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $managers = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $managers.ManagedBy.Add($i)
				                Write-Log -Status "DEBUG_MODE" -Message $managers.ManagedBy
				                #Changed following error action from SilentlyContinue to Stop in order to validade why the error handling is not working either when the cmdlet works or not	
                                Set-EXODistributionGroup $dl.Name -ManagedBy $managers.ManagedBy -ErrorAction Stop | Out-Null
                                
                                Write-Log -Status "GROUP_EDIT" -Message "The user $i was added sucessfully in the attribute ManagedBy of the group $groupName in the EXO."
				                Write-Log -Status "DEBUG_MODE" -Message "Set-EXODistributionGroup error handling didn't catch"
				                Write-Log -Status "DEBUG_MODE" -Message $managers.ManagedBy
                            }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the user $i in the attribute ManagedBy of the group $groupName in the EXO."
				                Write-Log -Status "DEBUG_MODE" -Message $_
				                Write-Log -Status "DEBUG_MODE" -Message "Set-EXODistributionGroup error handling catch"
				                Write-Log -Status "DEBUG_MODE" -Message $managers.ManagedBy
                            }
                        }
                    }
                }

                If($dl.ModeratedBy){
                    Set-EXODistributionGroup $dl.Name -ModerationEnabled $true | Out-Null

                    $arr = $dl.ModeratedBy.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $moderators = Get-EXODistributionGroup $dl.Name
                                $moderators.ModeratedBy.Add($i)
                                Set-EXODistributionGroup $dl.Name -ModeratedBy $moderators.ModeratedBy | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The user $i was added $i sucessfully in the attribute ModeratedBy of the group $groupName in the EXO"
                            }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the user $i in the attribute ModeratedBy of the group $groupName in the EXO"
                            }
                        }
                    }
                }

                If($dl.AcceptMessagesOnlyFromDLMembers){
                    $arr = $dl.AcceptMessagesOnlyFromDLMembers.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $acc1 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc1.AcceptMessagesOnlyFromDLMembers.Add($i)
                                Set-EXODistributionGroup $dl.Name -AcceptMessagesOnlyFromDLMembers $acc1.AcceptMessagesOnlyFromDLMembers -ErrorAction SilentlyContinue | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The group $i was added sucessfully in the attribute AcceptMessagesOnlyFromDLMembers of the group $groupName in the EXO"
                            }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the group $i in the attribute AcceptMessagesOnlyFromDLMembers of the group $groupName in the EXO"
                            }
                        }
                    }
                }
    
                If($dl.AcceptMessagesOnlyFrom){
                    $arr = $dl.AcceptMessagesOnlyFrom.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $acc2 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc2.AcceptMessagesOnlyFrom.Add($i)
                                Set-EXODistributionGroup $dl.Name -AcceptMessagesOnlyFrom $acc2.AcceptMessagesOnlyFrom -ErrorAction SilentlyContinue | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The user $i $i was added with sucess in the attribute AcceptMessagesOnlyFrom of the group $groupName"
                            }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the user $i in the attribute AcceptMessagesOnlyFrom of the group $groupName in the EXO"
                            }
                        }
                    }
                }

                If($dl.RejectMessagesFrom){
                    $arr = $dl.RejectMessagesFrom.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $acc3 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc3.RejectMessagesFrom.Add($i)
                                Set-EXODistributionGroup $dl.Name -RejectMessagesFrom $acc3.RejectMessagesFrom -ErrorAction SilentlyContinue | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The user $i was added sucessfully in the attribute RejectMessagesFrom of the group $groupName in the EXO"
                            }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the user $i in the attribute RejectMessagesFrom of the group $groupName in the EXO"
                            }
                        }
                    }
                }
 
                If($dl.RejectMessagesFromDLMembers){
                    $arr = $dl.RejectMessagesFromDLMembers.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $acc4 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc4.RejectMessagesFromDLMembers.Add($i)
                                Set-EXODistributionGroup $dl.Name -RejectMessagesFromDLMembers $acc4.RejectMessagesFromDLMembers -ErrorAction SilentlyContinue | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The group $i was added sucessfully in the attribute RejectMessagesFromDLMembers of the group $groupName in the EXO"
                            }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the group $i in the attribute RejectMessagesFromDLMembers of the group $groupName in the EXO"
                                }
                            }
                        
                        }
                    }

                If($dl.EmailAddresses){
                    $arr = $dl.EmailAddresses.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $acc5 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc5.EmailAddresses.Add($i)
                                $groupName = $dl.Name
                                Set-EXODistributionGroup $dl.Name -EmailAddresses $acc5.EmailAddresses -ErrorAction SilentlyContinue | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The email address $i was added sucessfully in the attribute EmailAddresses of the group $groupName in the EXO"
                                }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the email address $i in the attribute EmailAddresses of the group $groupName in the EXO"
                                }
                        }
                    }
                }

                If($dl.bypassModerationFromSendersOrMembers){
                    $arr = $dl.bypassModerationFromSendersOrMembers.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $acc6 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc6.bypassModerationFromSendersOrMembers.Add($i)
                                Set-EXODistributionGroup $dl.Name -bypassModerationFromSendersOrMembers $acc6.bypassModerationFromSendersOrMembers -ErrorAction SilentlyContinue | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The user/group $i was added sucessfully in the attribute bypassModerationFromSendersOrMembers of the group $groupName in the EXO"
                                }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the user/group $i in the attribute bypassModerationFromSendersOrMembers of the group $groupName in the EXO"
                                }
                        }
                    }
                }

                If($dl.GrantSendOnBehalfTo){
                    $arr = $dl.GrantSendOnBehalfTo.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                $acc7 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc7.GrantSendOnBehalfTo.Add($i)
                                Set-EXODistributionGroup $dl.Name -GrantSendOnBehalfTo $acc7.GrantSendOnBehalfTo -ErrorAction SilentlyContinue | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The user/group $i was added sucessfully in the attribute GrantSendOnBehalfTo of the group $groupName in the EXO"
                                }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to add the user/group $i in the attribute GrantSendOnBehalfTo of the group $groupName in the EXO"
                                }
                        }
                    }
                }
                                       
                If($dl.SendAsPermission){
                    $arr = $dl.SendAsPermission.Split(";")
                    foreach($i in $arr){
                        If($i){
                            try{
                                Add-EXORecipientPermission $dl.Name -AccessRights SendAs -Trustee $i -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                                Write-Log -Status "GROUP_EDIT" -Message "The permission SendAs was added sucessfully for the user/group $i in the group $groupName in the EXO"
                                }
                            catch{
                                Write-Log -Status "ERROR" -Message "Error when trying to the permission SendAs for the group $i in the group $groupName in the EXO"
                                }
                        }
                    }
                }
                              
                $dlsMembers = Import-Csv DistributionGroups_Members.csv | Where-Object{$_.DistributionGroup -eq $dl.Name}

                foreach($dlsMember in $dlsMembers){
                    
                    $dlMember = $dlsMember.Alias                    
                    try{
                        Add-EXODistributionGroupMember $dl.Name -Member $dlsMember.Alias -ErrorAction SilentlyContinue | Out-Null
                        Write-Log -Status "ADD_MEMBER" -Message "User $dlMember was added sucessfully in the members list of the group $groupname in the EXO"
                    }
                    catch{
                        Write-Log -Status "ERROR" -Message "Error when trying to add the user $dlMember in the members list of the group $groupname in the EXO"
                    }
                }           
            
            }
        }
        n{
            Write-Log -Status "ACTION" -Message "Creation of groups in the EXO aborted by user"

            Write-Host "Task aborted!"
        }
    }

    try{
        Remove-PSSession $EXOSession -ErrorAction SilentlyContinue | Out-Null
        Write-Log -Status "CONN" -Message "EXO Session removed sucessfully"
        }
    catch{
        Write-Log -Status "ERROR" -Message "Error when trying to remove the EXO session"
        }

    try{
        Remove-Item exo-contoso.sec -ErrorAction SilentlyContinue | Out-Null
        Write-Log -Status "CONN" -Message "File with the encripted credentials was removed sucessfully"
        }
    catch{
        Write-Log -Status "ERROR" -Message "Error when trying to remove the file with the encripted credentials"
    }
}