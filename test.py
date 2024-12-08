import re

line="this hdr-biz 123 model server 456"
pattern=r"123"
matchObj = re.search( pattern, line)
if(matchObj):
    print("ok")
else:
    print('fail')