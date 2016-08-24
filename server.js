"use strict";

var http = require("http");
var port = process.argv[2] || 8043;
var nodeStatic = require("node-static");

var file = new nodeStatic.Server("serverRoot", {
	cache: 3600,
	gzip: true
});

var server = http.createServer(function(request, response) {
	request.addListener("end", function() {
		file.serve(request, response);
	}).resume();
}).listen(port, function() {
	port = server.address().port;
	console.log("Listening on http://localhost:" + port);
});
