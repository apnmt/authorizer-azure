var http = require('https');

module.exports = function (context, req) {
    const objectId = (req.query.objectId || (req.body && req.body.objectId));
    context.log("Request Groups for User with userId=" + objectId);

    var body = "";
    body += 'grant_type=client_credentials';
    body += '&client_id=72dddf73-cc68-41af-b688-e7890045e72d';
    body += '&client_secret=A1T7Q~OlCVwyf49lF~cIdAyqvMelTVm05sHxw';
    body += '&scope=https://graph.microsoft.com/.default';

    const options = {
        hostname: 'login.microsoftonline.com',
        port: 443,
        path: '/d8e1390b-3103-4104-8237-628e88e6f47e/oauth2/v2.0/token',
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': body.length
        }
    }

    var response = '';
    const request = http.request(options, (res) => {
        context.log(`statusCode: ${res.statusCode}`)

        res.on('data', (d) => {
            response += d;
        })

        res.on('end', (d) => {
            var string = JSON.stringify(response);
            var json = JSON.parse(response);
            requestGroups(objectId, json['access_token'], function(data) {
                const jsonGroups = JSON.parse(data);
                const values = jsonGroups['value']
                var groups = '';
                values.forEach(function(v) {
                    groups += v['displayName']
                    groups += ';'
                });
                context.log("Successfully requested Groups for user=" + objectId + " groups=" + groups);
                context.res = {
                    body: {
                        "version": "1.0.0",
                        "action": "Continue",
                        groups
                    },
                    headers: {
                        "Content-Type": "application:json"
                    }
                }
                context.done();
            });
        })
    })

    request.on('error', (error) => {
        context.log.error(error)
        context.done();
    })

    request.write(body);

    function requestGroups(objectId, accessToken, callback) {
        var response = '';
        const options = {
            hostname: 'graph.microsoft.com',
            port: 443,
            path: '/v1.0/users/'+ objectId + '/memberOf',
            method: 'GET',
            headers: {
                "Authorization": "Bearer " + accessToken
            }
        }

        const request = http.request(options, res => {
            context.log(`statusCode: ${res.statusCode}`)

            res.on('data', (d) => {
                response += d;
            })

            res.on('end', (d) => {
                callback(response)
            });
        })
        request.on('error', error => {
            context.error(error)
        })
        request.write('data\n');
        request.write('data\n');
        request.end();
    }
};
