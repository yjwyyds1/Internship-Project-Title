// 零依赖静态文件服务器 — 无需 npm install，直接用 Node.js 内置模块
var http = require('http');
var fs = require('fs');
var path = require('path');

var PORT = 8080;
var ROOT = __dirname; // 项目根目录

var MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript',
  '.css':  'text/css',
  '.png':  'image/png',
  '.jpg':  'image/jpeg',
  '.gif':  'image/gif',
  '.json': 'application/json',
  '.b3dm': 'application/octet-stream',
  '.i3dm': 'application/octet-stream',
  '.pnts': 'application/octet-stream',
  '.cmpt': 'application/octet-stream',
  '.glb':  'model/gltf-binary',
  '.gltf': 'model/gltf+json',
  '.czml': 'application/json',
  '.kml':  'application/vnd.google-earth.kml+xml',
  '.kmz':  'application/vnd.google-earth.kmz',
};

var server = http.createServer(function (req, res) {
  var filePath = path.join(ROOT, decodeURIComponent(req.url.split('?')[0]));
  if (filePath.endsWith('/')) filePath = path.join(filePath, 'index.html');

  var ext = path.extname(filePath).toLowerCase();
  var contentType = MIME[ext] || 'application/octet-stream';

  fs.readFile(filePath, function (err, data) {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('404 Not Found: ' + req.url);
      return;
    }
    res.writeHead(200, {
      'Content-Type': contentType,
      'Access-Control-Allow-Origin': '*'
    });
    res.end(data);
  });
});

server.listen(PORT, function () {
  console.log('Server running at http://localhost:' + PORT + '/');
  console.log('  exp1: http://localhost:' + PORT + '/exp1_helloworld.html');
  console.log('  exp2: http://localhost:' + PORT + '/exp2_box_entity.html');
  console.log('Press Ctrl+C to stop.');
});
