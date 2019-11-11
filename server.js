/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var path = require('path');
var fs = require('fs');
var https = require('https');
var express = require('express');
var app = express();

// Set the address and the certificate.
var options = {
    hostname: 'localhost',
    key: fs.readFileSync('server.key'),
    cert: fs.readFileSync('server.crt'),
    ca: fs.readFileSync('ca.crt')
};

// Define the port. The service uses 'localhost' as the host address.
// Set the host member in the options object to set a custom host domain name or IP address.
var port = 8088;

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname + '/scripts'));

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname + '/media'));

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname + '/css'));

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname));

app.use(express.json())
// Set the route to the index.html file.
app.get('/', function (req, res) {
    var homepage = path.join(__dirname, 'index.html');
    res.sendFile(homepage);
});
var Client = require('node-rest-client').Client;
// direct way
var client = new Client();

app.get('/api', function (req, res) {
    console.log('Here');
    console.log(req.query.what);
    const abc = req.query.what;

    client.get("http://localhost:5000/a?what=" + abc + ' from node',
        function (data, response) {
            // parsed response body as js object
            console.log(data.toString());
            // raw response
        });
    console.log(req.query.what);
    res.sendStatus(200);
});

app.post('/api/p', function (req, res) {
    console.log('Here ');
    console.log(req.body);
    console.log(JSON.stringify(req.body, 0, 2));
    const abc = req.body.word;
    console.log('aa ' + req.body.word);
    var args = {
        data: req.body,
        headers: {"Content-Type": "application/json"}
    };
    const request = client.post("http://localhost:5000", args, function (data, response) {
        console.log('Herer ' + JSON.stringify(data, 0, 2));
        abcc = data;
        console.log('Prepare Sending');
        res.setHeader('Content-Type', 'application/json');
        res.send(JSON.stringify(data));
    });
    request.on('error', function (err) {
        console.log('request error', err);
        res.sendStatus(500);
    });

});

// Set the route for the HTML served to the dialog API call.
app.get('/dialogCount', function (req, res) {
    var homepage = path.join(__dirname, 'dialogCount.html');
    res.sendFile(homepage);
});

// Set the route for the HTML served to the dialog API call.
app.get('/dialogAlert', function (req, res) {
    var homepage = path.join(__dirname, 'dialogAlert.html');
    res.sendFile(homepage);
});

// Start the server.
https.createServer(options, app).listen(port, function () {
    console.log('Listening on https://localhost:' + port + '...');
});
