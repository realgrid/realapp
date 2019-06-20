var express = require('express');
var app     = express();
var path    = require('path');

app.use('/', express.static(__dirname + "/webapp"));

app.get('/', function (req, res) {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(4040, function () {
    console.log("Listening on port 4040");
});