
$sharepointSiteUrl = "https://m365x107527.sharepoint.com/sites/BuildingA-NT4/"
$clientID = "25e730c4-3859-4619-b1d5-bd41ffbd9635"
$clientSecret = ".eQSY8V/.=ZhpItaCczHa8nJyXZkh1c9"
$tenantName = "M365x107527.onmicrosoft.com"
$utcOffset = "-4"
$timezoneName = "Eastern Daylight Time"
$daysOutForMeetings = 21


$listName = "Meetings"
$now = (Get-Date).Date
$meetingExpiration = $now.AddDays($daysOutForMeetings)





