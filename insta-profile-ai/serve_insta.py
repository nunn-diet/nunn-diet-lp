import http.server, os
os.chdir('/Users/omorimaria/lp-preview/insta-profile-ai')
http.server.test(HandlerClass=http.server.SimpleHTTPRequestHandler, port=8769, bind='127.0.0.1')
