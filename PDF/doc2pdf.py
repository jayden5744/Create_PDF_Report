# -*- encoding: utf-8 -*-
# doc -> pdf 가능 https://qkqhxla1.tistory.com/402
import urllib2

# the input docx
file = open('demo1.docx', 'rb')
data = file.read()

#set up the request
req = urllib2.Request("http://converter-eval.plutext.com:80/v1/00000000-0000-0000-0000-000000000000/convert", data)
req.add_header('Content-Length', '%d' % len(data))
req.add_header('Content-Type', 'application/octet-stream')

# make the request
res = urllib2.urlopen(req)

# write the response to a file
pdf = res.read()
f = open('out.pdf', 'wb')
f.write(pdf)