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

- You must have or be able to consent to the following scopes for the enterprise app **Microsoft Graph Command Line Tools**

    - Agreement.Read.All
    - AdministrativeUnit.Read.All
    - Directory.Read.All
    - email
    - EntitlementManagement.Read.All
    - Group.Read.All
    - IdentityProvider.Read.All
    - openid
    - Organization.Read.All
    - PrivilegedAccess.Read.AzureAD
    - PrivilegedAccess.Read.AzureADGroup
    - PrivilegedAccess.Read.AzureResources
    - PrivilegedAssignmentSchedule.Read.AzureADGroup
    - PrivilegedEligibilitySchedule.Read.AzureADGroup
    - profile
    - RoleAssignmentSchedule.Read.Directory
    - RoleAssignmentSchedule.ReadWrite.Directory
    - RoleEligibilitySchedule.Read.Directory
    - RoleManagement.Read.All
    - RoleManagement.Read.Directory
    - RoleManagement.ReadWrite.Directory
    - RoleManagementAlert.Read.Directory
    - RoleManagementPolicy.Read.Directory
    - RoleManagementPolicy.Read.AzureADGroup
    - User.Read
    - User.Read.All
    - AgreementAcceptance.Read
    - AgreementAcceptance.Read.All
    - AuditLog.Read.All
    - Policy.Read.All

## Usage

```
.\PIMSCAN.ps1 -TenantId <TenantID> -Show -Verbose
```

Results are saved in a HTML file.

Open the Entra_ID_Role_Report_[TenantID].html in you did not supply the **-Show** parameter.

<br>