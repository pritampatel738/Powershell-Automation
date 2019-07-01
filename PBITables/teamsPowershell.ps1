# $uri = "https://outlook.office.com/webhook/ff84dff0-b7a1-4756-83a2-e087ffb6560e@e4d98dd2-9199-42e5-ba8b-da3e763ede2e/IncomingWebhook/005e6f799c6246deac6b0ee929c44faf/5ba00831-e4ad-457f-ba57-4a0e16d2e8bf"


# $body = ConvertTo-JSON @{
    # text = 'Hello Channel'
# }

# Invoke-RestMethod -uri $uri -Method Post -body $body -ContentType 'application/json'
connect-powerbiserviceaccount
#invoke-powertbirestmethod
$url = "https://api.powerbi.com/v1.0/myorg/datasets/f333789d-c173-4398-a9de-6baab1ba19d5/refreshes"
# get the latest refresh time using invoke-powerbirestmethod ....
#echo "Fetching refresh time from Workspace"
$refreshes = Invoke-PowerBIRestMethod -Url $url  -Method GET
echo $refreshes




























