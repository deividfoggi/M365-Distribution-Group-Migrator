#
#    Move-DistributionGroups.ps1
#
#    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
#    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,        
#    INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
#    We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute
#    the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks
#    to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on
#    Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us
#    and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the 
#    use or distribution of the Sample Code.
#    Please note: None of the conditions outlined in the disclaimer above will supersede the terms and conditions contained 
#    within the Premier Customer Services Description.
#
<#
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
You must follow the exection order, including the need to move to the EXO any mailboxes in the field ManagedBy and ModeratedBy before running the option 3

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
The execution of this script will create a log file in the same directory of this script with the name dd_MM_yyyy.LOG, creating a new log for each new day that the script is executed. This log will register actions and errors.
The file log contains the following columns separated by comma:

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

#Function to create a log file and register log entries
Function log{
    Param(
        [string]$Status,
        [string]$Message
    )
    
    $logName = Get-Date -Format d_M_yyyy.LOG

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

        log -Status "CONN" -Message "User credentials for the user $a.UserName encripted"

        $userName = $a.UserName
        $password = Get-Content .\exo-password.sec | convertto-securestring
        $EXOCredential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$password

        $Global:EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $EXOCredential -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue

        log -Status "CONN" -Message "Session with EXO using the user $userName created sucessfully"
    }
    catch{
        log -Status "ERROR" -Message "Error when trying to create a EXO session"
        }

    try{
        Import-PSSession $EXOSession -Prefix EXO -CommandName New-DistributionGroup,Add-DistributionGroupMember,Set-DistributionGroup,Get-DistributionGroup,Get-Mailbox,Add-RecipientPermission -AllowClobber -ErrorAction SilentlyContinue
        log -Status "CONN" -Message "Session with EXO imported sucessfully for the user $userName"
        }
        catch{
            log -Status "ERROR" -Message "Error when trying to import the EXO session"
        }
}

#Function to show the main menu
Function Menu{

    Write-Host "           
    Modo:

    1 - Export a list containing all the distribution groups and their respective members
    2 - Remove all groups from the Exchange OnPrem
    3 - Create all groups in the Exchange Online and add the respective members

    " -ForegroundColor Yellow

    $mode = Read-Host "Choose an execution mode"

    switch($mode){
        1{
            ExportGroups
        }
        2{
            RemoveGroups
        }
        3{
            CreateGroups
        }
    }
}

#Function to export groups
Function ExportGroups{

    Param(
        [string]$GroupList,
        [string]$Group
    )

    log -Status "ACTION" -Message "Option 1 - Export Groups and Members"

    If($GroupList -ne ''){
        try{
            $AllDG = Import-Csv $GroupList
            log -Status "ACTION" -Message "The option to migrate a list of groups from the file $GroupList was selected"
        }
        catch{
            log -Status "ERROR" -Message "Error when trying to import a list of groups"
        }
    }
    Else
    {
        If($Group -ne ''){
            try{
                $AllDG = Get-DistributionGroup $Group
                log -Status "ACTION" -Message "The option to migrate only the group $Group was selected"
            }
            catch{
                log -Status "ERROR" -Message "Error when trying to execute the cmdlet Get-DistributionGroup for the group $Group"
            }
        }
        Else
        {
            try{
            $AllDG = Get-DistributionGroup -resultsize unlimited | Select-Object Name,Alias,BypassNestedModerationEnabled,DisplayName,ManagedBy,MemberDepartRestriction,MemberJoinRestriction,ModeratedBy,ModerationEnabled,SendModerationNotifications,AcceptMessagesOnlyFromDLMembers,AcceptMessagesOnlyFrom,HiddenFromAddressListsEnabled,PrimarySmtpAddress,RejectMessagesFrom,RejectMessagesFromDLMembers,RequireSenderAuthenticationEnabled,EmailAddresses,bypassModerationFromSendersOrMembers,GrantSendOnBehalfTo,SendAsPermission
            log -Status "ACTION" -Message "The option to migrate all distribution groups was selected"
            }
            catch{
            log -Status "ERROR" -Message "Error when trying to executed the cmdlet Get-DistributionGroup for the group $Group"
            }
        }
    }

    $outputDL=@()
    $outputDLMembers=@()
    $arr=@()

    Foreach($obj in $AllDG){

        $AcceptMessagesOnlyFromDLMembers=""
        $AcceptMessagesOnlyFrom=""
        $RejectMessagesFrom=""
        $RejectMessagesFromDLMembers=""
        $EmailAddresses=""        
        $managedBy=""
        $moderatedBy=""
        $bypassModerationFromSendersOrMembers=""

        $objDL = New-Object psobject

        $objDL | Add-Member NoteProperty -Name "Name" -Value $obj.Name -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "Alias" -Value $obj.Alias -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "BypassNestedModerationEnabled" -Value $obj.BypassNestedModerationEnabled -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "DisplayName" -Value $obj.DisplayName -ErrorAction SilentlyContinue

        If($obj.ManagedBy -ne ''){
            $arr = $obj.ManagedBy | Select-Object Name
            foreach($user In $arr){
                $managedBy += $user.Name.ToString() + ";"        
            }
        }

        $objDL | Add-Member NoteProperty -Name "ManagedBy" -Value $managedBy -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "MemberDepartRestriction" -Value $obj.MemberDepartRestriction -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "MemberJoinRestriction" -Value $obj.MemberJoinRestriction -ErrorAction SilentlyContinue

        $arr=@()

        If($obj.ModeratedBy -ne ''){
            $arr = $obj.ModeratedBy | Select-Object Name
            foreach($user In $arr){
                $moderatedBy += $user.Name.ToString() + ";"        
            }
        }

        $objDL | Add-Member NoteProperty -Name "ModeratedBy" -Value $moderatedBy -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "ModerationEnabled" -Value $obj.ModerationEnabled -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "SendModerationNotifications" -Value $obj.SendModerationNotifications -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "MaxSendSize" -Value $obj.MaxSendSize -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "MaxReceiveSize" -Value $obj.MaxReceiveSize -ErrorAction SilentlyContinue

        $arr=@()

        If($obj.AcceptMessagesOnlyFromDLMembers -ne ''){
            $arr = $obj.AcceptMessagesOnlyFromDLMembers | Select-Object Name
            foreach($user In $arr){
                $AcceptMessagesOnlyFromDLMembers += $user.Name.ToString() + ";"        
            }
        }

        $objDL | Add-Member NoteProperty -Name "AcceptMessagesOnlyFromDLMembers" -Value $AcceptMessagesOnlyFromDLMembers -ErrorAction SilentlyContinue

        $arr=@()

        If($obj.AcceptMessagesOnlyFrom -ne ''){
            $arr = $obj.AcceptMessagesOnlyFrom | Select-Object Name
            foreach($user In $arr){
                $AcceptMessagesOnlyFrom += $user.Name.ToString() + ";"        
            }
        }

        $objDL | Add-Member NoteProperty -Name "AcceptMessagesOnlyFrom" -Value $AcceptMessagesOnlyFrom -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "HiddenFromAddressListsEnabled" -Value $obj.HiddenFromAddressListsEnabled -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "PrimarySmtpAddress" -Value $obj.PrimarySmtpAddress -ErrorAction SilentlyContinue

        $arr=@()

       If($obj.RejectMessagesFrom -ne ''){
            $arr = $obj.RejectMessagesFrom | Select-Object Name
            foreach($user In $arr){
                $RejectMessagesFrom += $user.Name.ToString() + ";"        
            }
        }
        
        $objDL | Add-Member NoteProperty -Name "RejectMessagesFrom" -Value $RejectMessagesFrom -ErrorAction SilentlyContinue

        $arr=@()

       If($obj.RejectMessagesFromDLMembers -ne ''){
            $arr = $obj.RejectMessagesFromDLMembers | Select-Object Name
            foreach($user In $arr){
                $RejectMessagesFromDLMembers += $user.Name.ToString() + ";"        
            }
        }
        
        $objDL | Add-Member NoteProperty -Name "RejectMessagesFromDLMembers" -Value $RejectMessagesFromDLMembers -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "RequireSenderAuthenticationEnabled" -Value $obj.RequireSenderAuthenticationEnabled -ErrorAction SilentlyContinue

        $arr=@()

        If($obj.EmailAddresses -ne ''){
            $arr = $obj.EmailAddresses | Select-Object SmtpAddress
            foreach($item In $arr){
                If($item -ne ''){
                    $EmailAddresses += $item.SmtpAddress + ";"
                }
            }
        }

        $arr=@()

       If($obj.bypassModerationFromSendersOrMembers -ne ''){
            $arr = $obj.bypassModerationFromSendersOrMembers | Select-Object Name
            foreach($user In $arr){
                $bypassModerationFromSendersOrMembers += $user.Name.ToString() + ";"        
            }
        }

        $arr=@()
        $GrantSendOnBehalfTo = ""

        If($obj.GrantSendOnBehalfTo -ne ''){
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

        $objDL | Add-Member NoteProperty -Name "EmailAddresses" -Value $EmailAddresses -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "bypassModerationFromSendersOrMembers" -Value $bypassModerationFromSendersOrMembers -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "GrantSendOnBehalfTo" -Value $GrantSendOnBehalfTo -ErrorAction SilentlyContinue
        $objDL | Add-Member NoteProperty -Name "SendAsPermission" -Value $SendAsPermission -ErrorAction SilentlyContinue

        $outputDL += $objDL
    
        $groupName = $objDL.Name
        log -Status "EXPORT_GROUP" -Message "The group $groupName was added to export list"

        }

        try{
            $outputDL | Export-Csv DistributionGroups.csv -NoTypeInformation -ErrorAction SilentlyContinue -Encoding UTF8
            log -Status "EXPORT_GROUP" -Message "The file DistributionGroups.csv was created sucessfully"
        }
        catch{
            log -Status "ERROR" -Message "Error when trying to create the file GruposDeDistribucao.csv"
        }
        
        Foreach($dg in $allDg){

            $Members = Get-DistributionGroupMember $dg.name -resultsize unlimited -ErrorAction SilentlyContinue
        
            Foreach($Member in $Members){
                $groupMember = $member.Alias

                $objMember = New-Object PSObject

                $objMember | Add-Member NoteProperty -Name "Alias" -Value $member.Alias -ErrorAction SilentlyContinue
                $objMember | Add-Member NoteProperty -Name "DistributionGroup" -Value $DG.Name -ErrorAction SilentlyContinue
        
                $outputDLMembers += $objMember

                log -Status "EXPORT_MEMBER" -Message "The member $groupMember of the group $groupName was added to the export list"
            }
        }

        try{
            $outputDLMembers | Export-Csv DistributionGroups_Members.csv -NoTypeInformation -ErrorAction SilentlyContinue -Encoding UTF8
            log -Status "EXPORT_GROUP" -Message "The file DistributionGroups_Members.csv was created with sucess"
        }
        catch{
            log -Status "ERROR" -Message "Error when trying to create the fileDistributionGroups_Members.csv"
        }
}

#Function used to remove groups from the Exchange OnPrem
Function RemoveGroups{

    Param(
        [string]$Group
    )
    
    log -Status "ACTION" -Message "Option 2 = Remove groups"

    $mode2 = Read-Host "
    WARNING!!! This procedure will remove all groups listed in the file DistributionGroups.csv from the On-Premises organization. Would you like to proceed (y/n)?"
    
    switch($mode2){
        y{
            log -Status "ACTION" -Message "The exclusion of all groups in the file DistributionGroups.csv was confirmed by the user"

            If($Group -ne ''){
                try{
                    $dls = Get-DistributionGroup $Group -ErrorAction SilentlyContinue
                    log -Status "ACTION" -Message "The option to migrate only the group $Group was selected"
                }
                catch{
                    log -Status "ERROR" -Message $varErro
                    $varErro = ''
                }
            }
            Else{
                $dls = Test-Path DistributionGroups.csv
                $dlsMembers = Test-Path DistributionGroups_Members.csv

                If($dls -eq $true -and $dlsMembers -eq $true){
                    try{
                        $dls = Import-Csv DistributionGroups.csv -ErrorAction SilentlyContinue
                        log -Status "ACTION" -Message "The option to migrate all groups was selected"
                    }
                    catch{
                        log -Status "ERROR" -Message $varErro
                        $varErro = ''
                    }
                }
                Else
                {
                    Write-Host "The file DistributionGroups.csv was removed. Keep this file in the same directory as the script or use the option 1 again to recreate them"
                    
                    log -Status "ERROR" -Message "File DistributionGroups.csv not found"
                }
            }
        
            $dls | ForEach-Object{
                $groupName = $_.Name
                
                try{
                    Remove-DistributionGroup $_.Name -BypassSecurityGroupManagerCheck -Confirm:$y -ErrorAction SilentlyContinue
                    log -Status "GROUP_REMOVED" -Message "Group $groupName was removed from the Exchange OnPrem"
                }
                catch{
                    log -Status "ERROR" -Message "Error when trying to remove the group $groupName"
                }
            }
        }
        n{
            log -Status "ACTION" -Message "Option 2 - Remove Groups aborted"
            Write-Host "Taks aborted!" -ForegroundColor Yellow
        }
    }
}

#Function to create groups in the EXO
Function CreateGroups{

    Param(
        [string]$Group
        )

    log -Status "ACTION" -Message "Option 3 - Create groups in the EXO was selected"

    $mode2 = Read-Host "
    WARNING!!! Before running the option 3, you must guarantee that the following steps have sucessfully completed:

        1. Only use option 3 after option 2;
        2. Move all users that are either in the field ManagedBy or ModeratedBy to the EXO;
        3. Run the sync, confirm that all groups have been removed from the Windows Azure Active Directory and make sure no error is logged.

    Proceed ONLY if all the 3 steps above have been sucessfully completed.

    Would you like to proceed (y/n)?"

    switch($mode2){
        y{
            log -Status "ACTION" -Message "The groups creation in the EXO was selected"

            $dls = Test-Path DistributionGroups.csv
            $dlsMembers = Test-Path DistributionGroups_Members.csv

            If($dls -eq $true -and $dlsMembers -eq $true){
                try{
                $dls = Import-Csv DistributionGroups.csv -ErrorAction SilentlyContinue
                log -Status "ACTION" -Message "The option to migrate all groups was selected"
            }
            catch{
                log -Status "ERROR" -Message "Error when trying to import the file DistributionGroups.csv"
                }
            }
            Else
            {
                Write-Host "The file DistributionGroups.csv was removed. Keep the file in the same directory as the script directory or use the option 2 to create them."
                    
                log -Status "ERROR" -Message "Arquivo DistributionGroups.csv não encontrado"
                
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
                    log -Status "GROUP_CREATED" -Message "Group $groupName in the EXO was created sucessfully"
                }
                catch{
                    log -Status "ERROR" -Message "Error when trying to create the group $groupName in the EXO"
                }
                
                try{
                    Set-EXODistributionGroup $dl.Name -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -PrimarySmtpAddress $dl.PrimarySmtpAddress -ErrorAction SilentlyContinue | Out-Null
                    log -Status "GROUP_EDIT" -Message "The attributes RequireSenderAuthenticationEnabled, HiddenFromAddressListsEnabled and PrimarySmtpAddress were changedto $RequireSenderAuthenticationEnabled, $HiddenFromAddressListsEnabled and $dl.PrimarySmtpAddress respectively in the group $groupName sucessfully in the EXO."
                }
                catch{
                    log -Status "ERROR" -Message "Error when trying to change the attributes RequireSenderAuthenticationEnabled, HiddenFromAddressListsEnabled and PrimarySmtpAddress in the grup $groupName in the EXO."
                }

                If($dl.ManagedBy -ne ''){
                    $arr = $dl.ManagedBy.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                $managers = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $managers.ManagedBy.Add($i)
				                log -Status "DEBUG_MODE" -Message $managers.ManagedBy
				                #Changed following error action from SilentlyContinue to Stop in order to validade why the error handling is not working either when the cmdlet works or not	
                                Set-EXODistributionGroup $dl.Name -ManagedBy $managers.ManagedBy -ErrorAction Stop | Out-Null
                                
                                log -Status "GROUP_EDIT" -Message "The user $i was added sucessfully in the attribute ManagedBy of the group $groupName in the EXO."
				                log -Status "DEBUG_MODE" -Message "Set-EXODistributionGroup error handling didn't catch"
				                log -Status "DEBUG_MODE" -Message $managers.ManagedBy
                            }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the user $i in the attribute ManagedBy of the group $groupName in the EXO."
				                log -Status "DEBUG_MODE" -Message $_
				                log -Status "DEBUG_MODE" -Message "Set-EXODistributionGroup error handling catch"
				                log -Status "DEBUG_MODE" -Message $managers.ManagedBy
                            }
                        }
                    }
                }

                If($dl.ModeratedBy -ne ''){
                    Set-EXODistributionGroup $dl.Name -ModerationEnabled $true | Out-Null

                    $arr = $dl.ModeratedBy.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                $moderators = Get-EXODistributionGroup $dl.Name
                                $moderators.ModeratedBy.Add($i)
                                Set-EXODistributionGroup $dl.Name -ModeratedBy $moderators.ModeratedBy | Out-Null
                                log -Status "GROUP_EDIT" -Message "The user $i was added $i sucessfully in the attribute ModeratedBy of the group $groupName in the EXO"
                            }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the user $i in the attribute ModeratedBy of the group $groupName in the EXO"
                            }
                        }
                    }
                }

                If($dl.AcceptMessagesOnlyFromDLMembers -ne ''){
                    $arr = $dl.AcceptMessagesOnlyFromDLMembers.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                $acc1 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc1.AcceptMessagesOnlyFromDLMembers.Add($i)
                                Set-EXODistributionGroup $dl.Name -AcceptMessagesOnlyFromDLMembers $acc1.AcceptMessagesOnlyFromDLMembers -ErrorAction SilentlyContinue | Out-Null
                                log -Status "GROUP_EDIT" -Message "The group $i was added sucessfully in the attribute AcceptMessagesOnlyFromDLMembers of the group $groupName in the EXO"
                            }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the group $i in the attribute AcceptMessagesOnlyFromDLMembers of the group $groupName in the EXO"
                            }
                        }
                    }
                }
    
                If($dl.AcceptMessagesOnlyFrom -ne ''){
                    $arr = $dl.AcceptMessagesOnlyFrom.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                $acc2 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc2.AcceptMessagesOnlyFrom.Add($i)
                                Set-EXODistributionGroup $dl.Name -AcceptMessagesOnlyFrom $acc2.AcceptMessagesOnlyFrom -ErrorAction SilentlyContinue | Out-Null
                                log -Status "GROUP_EDIT" -Message "The user $i $i was added with sucess in the attribute AcceptMessagesOnlyFrom of the group $groupName"
                            }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the user $i in the attribute AcceptMessagesOnlyFrom of the group $groupName in the EXO"
                            }
                        }
                    }
                }

                If($null -ne $dl.RejectMessagesFrom){
                    $arr = $dl.RejectMessagesFrom.Split(";")
                    foreach($i in $arr){
                        If($null -ne $i){
                            try{
                                $acc3 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc3.RejectMessagesFrom.Add($i)
                                Set-EXODistributionGroup $dl.Name -RejectMessagesFrom $acc3.RejectMessagesFrom -ErrorAction SilentlyContinue | Out-Null
                                log -Status "GROUP_EDIT" -Message "The user $i was added sucessfully in the attribute RejectMessagesFrom of the group $groupName in the EXO"
                            }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the user $i in the attribute RejectMessagesFrom of the group $groupName in the EXO"
                            }
                        }
                    }
                }
 
                If($dl.RejectMessagesFromDLMembers -ne ''){
                    $arr = $dl.RejectMessagesFromDLMembers.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                $acc4 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc4.RejectMessagesFromDLMembers.Add($i)
                                Set-EXODistributionGroup $dl.Name -RejectMessagesFromDLMembers $acc4.RejectMessagesFromDLMembers -ErrorAction SilentlyContinue | Out-Null
                                log -Status "GROUP_EDIT" -Message "The group $i was added sucessfully in the attribute RejectMessagesFromDLMembers of the group $groupName in the EXO"
                            }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the group $i in the attribute RejectMessagesFromDLMembers of the group $groupName in the EXO"
                                }
                            }
                        
                        }
                    }

                If($dl.EmailAddresses -ne ''){
                    $arr = $dl.EmailAddresses.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                $acc5 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc5.EmailAddresses.Add($i)
                                $groupName = $dl.Name
                                Set-EXODistributionGroup $dl.Name -EmailAddresses $acc5.EmailAddresses -ErrorAction SilentlyContinue | Out-Null
                                log -Status "GROUP_EDIT" -Message "The email address $i was added sucessfully in the attribute EmailAddresses of the group $groupName in the EXO"
                                }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the email address $i in the attribute EmailAddresses of the group $groupName in the EXO"
                                }
                        }
                    }
                }

                If($dl.bypassModerationFromSendersOrMembers -ne ''){
                    $arr = $dl.bypassModerationFromSendersOrMembers.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                $acc6 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc6.bypassModerationFromSendersOrMembers.Add($i)
                                Set-EXODistributionGroup $dl.Name -bypassModerationFromSendersOrMembers $acc6.bypassModerationFromSendersOrMembers -ErrorAction SilentlyContinue | Out-Null
                                log -Status "GROUP_EDIT" -Message "The user/group $i was added sucessfully in the attribute bypassModerationFromSendersOrMembers of the group $groupName in the EXO"
                                }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the user/group $i in the attribute bypassModerationFromSendersOrMembers of the group $groupName in the EXO"
                                }
                        }
                    }
                }

                If($dl.GrantSendOnBehalfTo -ne ''){
                    $arr = $dl.GrantSendOnBehalfTo.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                $acc7 = Get-EXODistributionGroup $dl.Name -ErrorAction SilentlyContinue
                                $acc7.GrantSendOnBehalfTo.Add($i)
                                Set-EXODistributionGroup $dl.Name -GrantSendOnBehalfTo $acc7.GrantSendOnBehalfTo -ErrorAction SilentlyContinue | Out-Null
                                log -Status "GROUP_EDIT" -Message "The user/group $i was added sucessfully in the attribute GrantSendOnBehalfTo of the group $groupName in the EXO"
                                }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to add the user/group $i in the attribute GrantSendOnBehalfTo of the group $groupName in the EXO"
                                }
                        }
                    }
                }
                                       
                If($dl.SendAsPermission -ne ''){
                    $arr = $dl.SendAsPermission.Split(";")
                    foreach($i in $arr){
                        If($i -ne ''){
                            try{
                                Add-EXORecipientPermission $dl.Name -AccessRights SendAs -Trustee $i -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                                log -Status "GROUP_EDIT" -Message "The permission SendAs was added sucessfully for the user/group $i in the group $groupName in the EXO"
                                }
                            catch{
                                log -Status "ERROR" -Message "Error when trying to the permission SendAs for the group $i in the group $groupName in the EXO"
                                }
                        }
                    }
                }
                              
                $dlsMembers = Import-Csv DistributionGroups_Members.csv | Where-Object{$_.DistributionGroup -eq $dl.Name}

                foreach($dlsMember in $dlsMembers){
                    
                    $dlMember = $dlsMember.Alias                    
                    try{
                        Add-EXODistributionGroupMember $dl.Name -Member $dlsMember.Alias -ErrorAction SilentlyContinue | Out-Null
                        log -Status "ADD_MEMBER" -Message "User $dlMember was added sucessfully in the members list of the group $groupname in the EXO"
                    }
                    catch{
                        log -Status "ERROR" -Message "Error when trying to add the user $dlMember in the members list of the group $groupname in the EXO"
                    }
                }           
            
            }
        }
        n{
            log -Status "ACTION" -Message "Creation of groups in the EXO aborted by user"

            Write-Host "Task aborted!"
        }
    }

    try{
        Remove-PSSession $EXOSession -ErrorAction SilentlyContinue | Out-Null
        log -Status "CONN" -Message "EXO Session removed sucessfully"
        }
    catch{
        log -Status "ERROR" -Message "Error when trying to remove the EXO session"
        }

    try{
        Remove-Item exo-password.sec -ErrorAction SilentlyContinue | Out-Null
        log -Status "CONN" -Message "File with the encripted credentials was removed sucessfully"
        }
    catch{
        log -Status "ERROR" -Message "Error when trying to remove the file with the encripted credentials"
    }
}

If($Group -eq ''){
    Menu
}
Else{
    If($Group -ne ''){
        ExportGroups -Group $Group
        RemoveGroups -Group $Group
        CreateGroups
        }
    }