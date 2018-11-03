//var AuthenticationContext = require('adal-node').AuthenticationContext;
var request = require('request-promise');
var spauth = require('node-sp-auth');
var siteUrl = 'https://cosmo2013.sharepoint.com/sites/DigitalPlace';
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
        var restApi = "https://cosmo2013.sharepoint.com/sites/DigitalPlace/_api/web/lists/getbytitle('Topics')/items?$select=Title,FileRef,IsExternal,TopicImage,Body&$filter=(IsExternal eq 0)";
        request.get({
            url: restApi,
            headers: data.headers,
            json: true
        }).then(function (listData) {
            var listResults = listData.d.results;
            context.res = {
                // status: 200, /* Defaults to 200 */
                body: listResults.length
            };    
        });
    /*if (req.query.name || (req.body && req.body.name)) {
        context.res = {
            // status: 200, 
            body: "Hello, " + (req.query.name || req.body.name)
        };
    }
    else {
        context.res = {
            status: 400,
            body: "Please pass a name on the query string or in the request body"
        };
    }*/
    });
};