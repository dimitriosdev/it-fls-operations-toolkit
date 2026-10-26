Connect-MgGraph -Scopes "ChannelMessage.Read.All"

Get-MgTeamChannelMessage `
    -TeamId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ChannelId "19:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@thread.tacv2" `
    -Top 5 |
    Select-Object Id, CreatedDateTime