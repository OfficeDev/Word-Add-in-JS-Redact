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


// Set the route to the index.html file.
app.get('/', function(req, res) {
    var homepage = path.join(__dirname, 'index.html');
    res.sendFile(homepage);
});

// Set the route for the HTML served to the dialog API call.
app.get('/dialogCount', function(req, res) {
    var homepage = path.join(__dirname, 'dialogCount.html');
    res.sendFile(homepage);
});

// Set the route for the HTML served to the dialog API call.
app.get('/dialogAlert', function(req, res) {
    var homepage = path.join(__dirname, 'dialogAlert.html');
    res.sendFile(homepage);
});

// Start the server.
https.createServer(options, app).listen(port, function() {
    console.log('Listening on https://localhost:' + port + '...');
});
