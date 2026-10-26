# Teams Channel Export - Validation Test

## Purpose

Before running a full export, this validation test confirms that:

1. Microsoft Graph PowerShell is installed
2. Authentication works
3. Graph permissions are sufficient
4. The target Team can be accessed
5. The target Channel can be read

If this test succeeds, the full export has a high probability of succeeding as well.

---

## Prerequisites

### Install Microsoft Graph

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Connect to Microsoft Graph

```powershell
Connect-MgGraph -Scopes "ChannelMessage.Read.All"
```

A browser window will open and you must sign in using the account that has access to the Team.

---

## Validation Test

### Team Information

```text
Team Name : My team name
Team Id   : xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

Channel Id: 19:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@thread.tacv2
```

### Verify Authentication

```powershell
Get-MgContext
```

Expected:

```text
Scopes:
ChannelMessage.Read.All
```

---

### Verify Team Access

```powershell
Get-MgTeam `
    -TeamId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

Expected:

```text
DisplayName : IT-FLS Wohlen
```

---

### Verify Message Access

Retrieve the first 5 messages from the channel:

```powershell
Get-MgTeamChannelMessage `
    -TeamId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ChannelId "19:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@thread.tacv2" `
    -Top 5
```

Expected:

A collection of Teams messages should be returned.

Example:

```text
Id                  : 1701234567890
CreatedDateTime     : 2026-10-01T07:21:11Z
```

---

## Common Errors

### Need Admin Approval

```text
Need admin approval
```

Meaning:

The account has not been granted consent for:

```text
ChannelMessage.Read.All
```

---

### Access Denied

```text
403 Forbidden
```

Meaning:

The account lacks permission to read Teams messages.

---

### Channel Not Found

```text
404 Not Found
```

Meaning:

Incorrect Team ID or Channel ID.

---

## Success Criteria

The validation is considered successful if:

- Authentication succeeds
- Team information is returned
- At least one channel message is returned

Once validated, proceed with:

```text
Export-TeamsChannel.ps1
```

to export channel conversations and replies into JSON.