# Move-DistributionGroups.ps1

**.SYNOPSIS**
Export all groups/members to CSV files e after creates them in the EXO, finally adds each member in their respective groups
Move-DistributionGroups.ps1

**.DESCRIPTION**
The Move-DistributionGroups script exports the attributes and members of all distribution groups from the OnPrem to the EXO, moreover adds the members in their respective groups

Obs: You must execute this script with a credential that has acess to all groups onprem and also you should has a credential member of the role needed to create and edit EXO distribution groups

**.ATTRIBUTES**
The following attributes will be copied

    Name,Alias,BypassNestedModerationEnabled,DisplayName,ManagedBy,MemberDepartRestriction,MemberJoinRestriction,ModeratedBy,ModerationEnabled,SendModerationNotifications,AcceptMessagesOnlyFromDLMembers,AcceptMessagesOnlyFrom,HiddenFromAddressListsEnabled,PrimarySmtpAddress,RejectMessagesFrom,RejectMessagesFromDLMembers,RequireSenderAuthenticationEnabled,EmailAddresses,bypassModerationFromSendersOrMembers,GrantSendOnBehalfTo,SendAsPermission

**.PreREQUISITES**
You must follow the exection order, including the need to move to the EXO any mailboxes in the field ManagedBy and ModeratedBy before running the option 3

**.PARAMETER Available Options**
1 - Export a list containing all the distribution groups and their respective members

The first option exports a list containing all distribution groups to the file DistributionGroups.csv. Also, exports a list of members per group to the file DistributionGroups_Members.csv. This option must be the first one to be executed in the case
of a cut over migration. Any files with the same name in the script directory will be overwritten.

2 - Remove all groups from the Exchange OnPrem

The second option removes all groups listed in the DistributionGroups.csv from the Exchange OnPrem. It is mandatory to run the first option before the second. Is necessary to have the file DistributionGroups.csv in the same path as the script path.

3 - Create all groups in the Exchange Online and add the respective members

The third option creates all groups listed in the file DistributionGroups.csv into EXO, edits them with the attributes previously exported to the file DistributionGroups.csv

**.PARAMETER Group**

The Group parameter could be used to move a specific group

**.LOGGING**
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
