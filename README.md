<p>
  <h1 align="center"><b>⚔ <img src="https://github.com/canix1/PIMSCAN/blob/main/img/PIMSCAN.png" width="70%"> ⚔</b></h1>
</p>
<p>
<h2 align="center"><b>A tool to create reports on Entra ID Role Assignments.</b></h2>
</p>
<p>
  <h2 align="center"><img src="https://github.com/canix1/PIMSCAN/blob/main/img/Animation.gif" width="80%"> </h2>
</p>

## Prerequisites

- PowerShell Module: MSAL.PS
```
Install-module MSAL.PS -Scope CurrentUser -Force -Confirm:$False
```

### Minumum Permissions with limited data

- Use the parameter **-LimitedReadOnly**, .\PIMSCAN.ps1 -TenantId [Tenant ID] -Show -verbose **-LimitedReadOnly**

- Global Reader role

- Consent for these:
  - AdministrativeUnit.Read.All
  - Directory.Read.All
  - Group.Read.All
  - PrivilegedAccess.Read.AzureAD
  - PrivilegedAccess.Read.AzureADGroup
  - PrivilegedAccess.Read.AzureResources
  - PrivilegedAssignmentSchedule.Read.AzureADGroup
  - PrivilegedEligibilitySchedule.Read.AzureADGroup
  - RoleAssignmentSchedule.Read.Directory
  - RoleEligibilitySchedule.Read.Directory
  - RoleManagement.Read.All
  - RoleManagement.Read.Directory
  - RoleManagementAlert.Read.Directory
  - RoleManagementPolicy.Read.Directory
  - RoleManagementPolicy.Read.AzureADGroup
  - User.Read
  - User.Read.All
  - offline_access


Run the following grant command as a Global Admin to grant a specific user the read-only scopes.
```
Install-Module Microsoft.Graph -Scope CurrentUser

connect-MgGraph -Scopes "Directory.AccessAsUser.All" -TenantId "<Your Tenant ID>"

$scopesOnlyRead = "AdministrativeUnit.Read.All Directory.Read.All Group.Read.All PrivilegedAccess.Read.AzureAD PrivilegedAccess.Read.AzureADGroup PrivilegedAccess.Read.AzureResources PrivilegedAssignmentSchedule.Read.AzureADGroup PrivilegedEligibilitySchedule.Read.AzureADGroup RoleAssignmentSchedule.Read.Directory RoleEligibilitySchedule.Read.Directory RoleManagement.Read.All RoleManagement.Read.Directory RoleManagementAlert.Read.Directory RoleManagementPolicy.Read.Directory RoleManagementPolicy.Read.AzureADGroup User.Read User.Read.All offline_access"
$params = @{
     # Microsoft Graph Command Line Tools
     ClientId = "4ad243ae-ea7f-4496-949e-4c64f1e96d71"
     # Singe User Consent
     ConsentType = "Principal"
     # Prinicpal to allow consent for
     PrincipalId = "<Prinicipal Object ID>"
     # GraphAggregatorService
     ResourceId = "4131d640-34dd-4690-ad11-45ddcd773304"
     # List of scopes/permissions
     Scope =  $scopesOnlyRead
}

New-MgOauth2PermissionGrant -BodyParameter $params
```

You will not be able to collect the data in the table below with Read-Only

|Object|Attribute|Description|Required Permission|
| -------- | ------- | -------- | -------- | 
|roleAssignmentScheduleRequests|justification|Supplied justification|RoleEligibilitySchedule.ReadWrite.Directory|
|roleAssignmentScheduleRequests|status|State of the request|RoleEligibilitySchedule.ReadWrite.Directory|
|roleAssignmentScheduleRequests|createdDateTime|Creation date of the request|RoleEligibilitySchedule.ReadWrite.Directory|
|roleEligibilityScheduleRequests|justification|Supplied justification|RoleEligibilitySchedule.ReadWrite.Directory|
|roleEligibilityScheduleRequests|status|State of the request|RoleEligibilitySchedule.ReadWrite.Directory|
|roleEligibilityScheduleRequests|createdDateTime|Creation date of the request|RoleEligibilitySchedule.ReadWrite.Directory|

### Full access with Write scopes for roleAssignmentScheduleRequests and roleEligibilityScheduleRequests.
- You must have or be able to consent to the following scopes for the enterprise app **Microsoft Graph Command Line Tools**

  - AdministrativeUnit.Read.All
  - Directory.Read.All
  - Group.Read.All
  - PrivilegedAccess.Read.AzureAD
  - PrivilegedAccess.Read.AzureADGroup
  - PrivilegedAccess.Read.AzureResources
  - PrivilegedAssignmentSchedule.Read.AzureADGroup
  - PrivilegedEligibilitySchedule.Read.AzureADGroup
  - RoleAssignmentSchedule.Read.Directory
  - RoleAssignmentSchedule.ReadWrite.Directory
  - RoleEligibilitySchedule.Read.Directory
  - RoleEligibilitySchedule.ReadWrite.Directory
  - RoleManagement.Read.All
  - RoleManagement.Read.Directory
  - RoleManagementAlert.Read.Directory
  - RoleManagementPolicy.Read.Directory
  - RoleManagementPolicy.Read.AzureADGroup
  - User.Read
  - User.Read.All
  - offline_access

Run the following grant command as a Global Admin to grant a specific user the read-only scopes.
```
Install-Module Microsoft.Graph -Scope CurrentUser

connect-MgGraph -Scopes "Directory.AccessAsUser.All" -TenantId "<Your Tenant ID>"

$scopesWrite = "AdministrativeUnit.Read.All Directory.Read.All Group.Read.All PrivilegedAccess.Read.AzureAD PrivilegedAccess.Read.AzureADGroup PrivilegedAccess.Read.AzureResources PrivilegedAssignmentSchedule.Read.AzureADGroup PrivilegedEligibilitySchedule.Read.AzureADGroup RoleAssignmentSchedule.Read.Directory RoleAssignmentSchedule.ReadWrite.Directory RoleEligibilitySchedule.Read.Directory RoleEligibilitySchedule.ReadWrite.Directory RoleManagement.Read.All RoleManagement.Read.Directory RoleManagementAlert.Read.Directory RoleManagementPolicy.Read.Directory RoleManagementPolicy.Read.AzureADGroup User.Read User.Read.All offline_access"

$params = @{
     # Microsoft Graph Command Line Tools
     ClientId = "4ad243ae-ea7f-4496-949e-4c64f1e96d71"
     # Singe User Consent
     ConsentType = "Principal"
     # Prinicpal to allow consent for
     PrincipalId = "<Prinicipal Object ID>"
     # GraphAggregatorService
     ResourceId = "4131d640-34dd-4690-ad11-45ddcd773304"
     # List of scopes/permissions
     Scope =  $scopesWrite
}

New-MgOauth2PermissionGrant -BodyParameter $params
```

## Usage

### Read-Only Limited

```
.\PIMSCAN.ps1 -TenantId <TenantID> -Show -Verbose -LimitedReadOnly
```
### Get all data
```
.\PIMSCAN.ps1 -TenantId <TenantID> -Show -Verbose
```



Results are saved in a HTML file.

Open the Entra_ID_Role_Report_[TenantID].html if you did not used the **-Show** parameter.

<br>