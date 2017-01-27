const adal = require('adal-node');
const request = require('request');

const TENANT = "tenant.onmicrosoft.com";
const GRAPH_URL = "https://graph.microsoft.com";
const CLIENT_ID = "copy and paste your client id";
const CLIENT_SECRET = "create a secret key and add it here";
const GROUP_ID = "enter-the-group-id";

function getToken() {
    return new Promise((resolve, reject) => {
        const authContext = new adal.AuthenticationContext(`https://login.microsoftonline.com/${TENANT}`);
        authContext.acquireTokenWithClientCredentials(GRAPH_URL, CLIENT_ID, CLIENT_SECRET, (err, tokenRes) => {
            if (err) {
                reject(err);
            }
            var accesstoken = tokenRes.accessToken;
            resolve(accesstoken);
        });
    });
}


getToken().then(token => {
    /* Get the group details */
    let options = {
        method: 'GET',
        url: `https://graph.microsoft.com/v1.0/groups/${GROUP_ID}`,
        headers: {
            'Authorization': 'Bearer ' + token,
            'content-type': 'application/json'
        }
    };

    request(options, (error, response, body) => {
        if (!error && response.statusCode == 200) {
            var result = JSON.parse(body);
            // Log all the keys and values of the group
            for (var key in result) {
                console.log(`${key}: ${JSON.stringify(result[key])}`);
            }

            /* Get the url of the group site */
            options.url = `https://graph.microsoft.com/v1.0/groups/${GROUP_ID}/drive/root/webUrl`

            request(options, (error, response, body) => {
                if (!error && response.statusCode == 200) {
                    const result = JSON.parse(body);
                    const webUrl = result.value;
                    console.log(`webUrl: ${webUrl.substring(0, webUrl.lastIndexOf('/'))}`);
                }
            });
        }
    });
});