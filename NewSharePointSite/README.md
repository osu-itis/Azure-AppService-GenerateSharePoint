# NewSharePointSite

## About

Generate a new SharePoint site

## Use

### Post Request

```HTTP REST
POST /api/NewSharePointSite?code=<AZUREFUNCTIONKEY>
Host: <HOST>
Content-Type: application/json

{
    "ticketID":"00000000",
    "owner":"username@oregonstate.edu",
    "displayName":"Sharepoint Name",
    "description":"Sharepoint Description"
}
```

#### Response

```JSON
{
    "id": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
    "displayName": "Sharepoint Name",
    "description": "Sharepoint Description",
    "webUrl": "https://oregonstateuniversity.sharepoint.com/sites/SharepointName",
    "mail": "SharepointName@OregonStateUniversity.onmicrosoft.com",
    "mailNickname": "SharepointName",
    "visibility": "Private",
    "createdDateTime": "0000-00-00T00:00:00Z"
}
```
