var request = require('request-promise');
var spauth = require('node-sp-auth');
var siteUrl = 'https://nakkeerann.sharepoint.com';
var clientId = 'ac75fede-95f5-4c00-9256-f163d1beb46a';
var clientSecret = '8T3dEmEcqu2U5pYpwwgfhwECk7vc0ybszB7bUaPQ9qI=';
module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');
    spauth.getAuth(siteUrl, {
        clientId: clientId,
        clientSecret: clientSecret
    }).then(function (data) {
        context.res = {
            status: 200,
            body: data
        };
        var restApi = "https://nakkeerann.sharepoint.com/_api/web/lists/getbytitle('Books')/items?$select=Title";
        request.get({
            url: restApi,
            headers: data.headers,
            json: true
        }).then(function (listData) {
            var listResults = listData.value;
            context.res = {
                // status: 200, /* Defaults to 200 */
                body: listResults
            };    
        });
    });
};