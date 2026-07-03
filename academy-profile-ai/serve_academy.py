import http.server, os
os.chdir('/Volumes/SSD-PHPU3A/クラウドコード真里亜/academy-profile-ai')
http.server.HTTPServer(('', 8770), http.server.SimpleHTTPRequestHandler).serve_forever()
