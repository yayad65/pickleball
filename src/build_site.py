import re
import os
HERE=os.path.dirname(os.path.abspath(__file__)); OUT=os.path.dirname(HERE)
src=open(os.path.join(HERE,'pickleball.html'),encoding='utf-8').read()
FONTS_DINK='https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&family=Dancing+Script:wght@700&display=swap'
def dinkify(s):
    s=s.replace('<title>Pickleball Round Robin</title>','<title>Dink &amp; Drive</title>\n<script>window.__THEME=\'dink\';document.documentElement.classList.add(\'dink\');</script>',1)
    s=re.sub(r'https://fonts.googleapis.com/css2\?family=Atkinson[^"]+',FONTS_DINK,s,count=1)
    return s
FONTS_PAPER='https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800;900&family=Barlow+Condensed:wght@300;400&display=swap'
def paperify(s):
    s=s.replace('<title>Pickleball Round Robin</title>','<title>Dink &amp; Drive</title>\n<script>window.__THEME=\'paper\';document.documentElement.classList.add(\'paper\');</script>',1)
    s=re.sub(r'https://fonts.googleapis.com/css2\?family=Atkinson[^"]+',FONTS_PAPER,s,count=1)
    return s
def site(s,title,color):
    head=f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="{title}">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="theme-color" content="{color}">
<link rel="apple-touch-icon" href="apple-touch-icon.png">
<link rel="manifest" href="manifest.webmanifest">
<style>:root{{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}}body{{margin:0}}img{{max-width:100%}}[hidden]{{display:none!important}}</style>
'''
    s=s.replace('<meta name="claude:viewport-fit-sides" content="cover">\n','')
    i=s.index('</style>')+len('</style>')
    return head+s[:i]+'\n</head>\n<body>\n'+s[i:]+'\n</body>\n</html>\n'
open(os.path.join(OUT,'classic.html'),'w',encoding='utf-8').write(site(src,'Pickleball','#1559B0'))
open(os.path.join(OUT,'index.html'),'w',encoding='utf-8').write(site(paperify(src),'Dink &amp; Drive','#EEF0F4'))
open(os.path.join(OUT,'dink.html'),'w',encoding='utf-8').write(site(dinkify(src),'Dink &amp; Drive','#3D8279'))
open(os.path.join(OUT,'paper.html'),'w',encoding='utf-8').write(site(paperify(src),'Dink &amp; Drive','#EEF0F4'))
print('built')
