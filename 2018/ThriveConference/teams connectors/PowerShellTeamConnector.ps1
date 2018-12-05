
# In a Teams Channel click on the three dots "..." next to the channel name on top of the window
# Click on connectors
# create a generic webhook connector
# paste the connector url you created into the $uri variables to get startet


# 01 - Simple connector with just a random text
$uri = '{YOUR_CHANNEL_CONNECTOR_URL}'
$body = ConvertTo-JSON @{
    text = 'Dober dan! 123'
}

Invoke-RestMethod -uri $uri -Method Post -body $body -ContentType 'application/json'



# 02 - Advanced connector with a activity including image, text and a button
$uri = '{YOUR_CHANNEL_CONNECTOR_URL}'
# these values would be retrieved from or set by an application
$status = 'success'
$fact1 = '[Ljubljana](https://en.wikipedia.org/wiki/Ljubljana)'
$fact2 = '[More Infos](https://www.visitljubljana.com/en/visitors/)'

$body = ConvertTo-Json -Depth 4 @{
    title = 'Hello Ljubljana'
    text = "This demo is a $status"
    sections = @(
        @{
            activityTitle = 'Infos'
            activitySubtitle = 'You want to know more?'
            activityText = 'Ljubljana is the capital and largest city of Slovenia. It has been the cultural, educational, economic, political, and administrative center of independent Slovenia since 1991.'
            activityImage = 'https://i.ytimg.com/vi/V56Fmj2juBw/maxresdefault.jpg' # this value would be a path to a nice image you would like to display in notifications
        },
        @{
            title = 'Details'
            facts = @(
                @{
                name = 'Wikipedia'
                value = $fact1
                },
                @{
                name = 'Visitor Info'
                value = $fact2
                }
            )
        }
    )
    potentialAction = @(@{
            '@context' = 'http://schema.org'
            '@type' = 'ViewAction'
            name = 'Click here to visit PowerShell.org'
            target = @('http://powershell.org')
        })
}


Invoke-RestMethod -uri $uri -Method Post -body $body -ContentType 'application/json'