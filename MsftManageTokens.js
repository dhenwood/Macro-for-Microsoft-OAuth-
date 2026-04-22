import xapi from 'xapi';

const CLIENT_ID = "update"
const CLIENT_SECRET = "update"
const TENANT_ID = "update"
const SAVED_TOKEN_FILE = "MsftSavedTokens"

async function getStoredTokens() {
  var macro = ''
  try {
    macro = await xapi.Command.Macros.Macro.Get({
      Name: SAVED_TOKEN_FILE,
      Content: 'True'
    })
  } catch (error) {
    return error
  }
  let raw = macro.Macro[0].Content.replace(/var.*memory.*=\s*{/g, '{')
  let data = JSON.parse(raw)
  return data
}

async function saveTokens(tokens) {
  let newStore = JSON.stringify(tokens, null, 4);
  xapi.Command.Macros.Macro.Save({
    Name: SAVED_TOKEN_FILE
  }, `var memory = ${newStore}`)
}


async function refreshToken(){
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`
  const header = "Content-type: application/x-www-form-urlencoded"
  const body = `client_id=${CLIENT_ID}&client_secret=${CLIENT_SECRET}&scope=https://graph.microsoft.com/.default&grant_type=client_credentials`
  try {
    const response = await xapi.Command.HttpClient.Post({
        AllowInsecureHTTPS: "True",
        Header: header,
        ResultBody: "PlainText",
        Url: url
      },
      body);

    let responseBody = response.Body
    let data = JSON.parse(responseBody)
    if (!data.access_token) {
      throw new Error('Invalid refresh response');
    }
    const now = Date.now();

    const tokens = {
      access_token: data.access_token,
      expires_at: now + (data.expires_in * 1000)
    };

    await saveTokens(tokens);
    return tokens
  } catch (err) {
    console.error('Refresh failed:', err);
    return null
  }
}

async function getValidAccessToken() {
  const tokens = await getStoredTokens();

  if (!tokens) {
    throw new Error('No tokens stored');
  }

  const now = Date.now();
  // Refresh 1 min before expiry
  if (now > tokens.expires_at - 60000) {
    console.log('Access token expired or about to expire, refreshing...');
    const newTokens = await refreshToken();
    return newTokens.access_token;
  }else{
    console.log("current token valid")
    return tokens.access_token;
  }
}

export{getValidAccessToken}
