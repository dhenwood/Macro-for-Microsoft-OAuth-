import xapi from 'xapi';
import {getValidAccessToken} from './MsftManageTokens';

async function listAllUsers(){
  let url = "https://graph.microsoft.com/v1.0/users"
  let token = await getValidAccessToken()
  let header = `Authorization: Bearer ${token}`

  try {
    const response = await xapi.Command.HttpClient.Get({
        AllowInsecureHTTPS: "True",
        Header: header,
        ResultBody: "PlainText",
        Url: url
    });

    let responseBody = response.Body
    let data = JSON.parse(responseBody)
    if (!data.value) {
      throw new Error('Invalid data response');
    }else{
      let items = data.value
      items.forEach((item) => {
        let displayName = item.displayName
        console.log("displayName: " + displayName)
      })
    }
  } catch (err) {
    console.error('Refresh failed:', err);
  }
}

listAllUsers()
