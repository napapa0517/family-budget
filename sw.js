// 설치형 웹앱용 최소 서비스 워커. 데이터는 저장하지 않고 그대로 인터넷에서 가져옴.
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function () {});
