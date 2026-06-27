import http.server, os, sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

class Handler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a): pass

http.server.HTTPServer(('', 8771), Handler).serve_forever()
