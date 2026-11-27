"""Builds the Bar Card.
  python build_card.py           -> ../../index.html (the page GitHub Pages serves)
Combines bar_card_template.html (the page) with data.json (stats, schedule, rosters, settings).
"""
import json, os
HERE=os.path.dirname(os.path.abspath(__file__))
t=open(os.path.join(HERE,'bar_card_template.html'),encoding='utf-8').read()
d=json.load(open(os.path.join(HERE,'data.json'),encoding='utf-8'))
body=t.replace('/*DATA*/',json.dumps(d,separators=(',',':'),ensure_ascii=False)).replace('<meta charset="utf-8">\n','',1)
body=body.replace('<meta name="claude:viewport-fit-sides" content="cover">\n','')
i=body.index('<main>')
head=('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
 '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
 '<style>:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0}img{max-width:100%}</style>\n')
out=os.path.join(HERE,'..','..','index.html')
open(out,'w',encoding='utf-8').write(head+body[:i]+'</head>\n<body>\n'+body[i:]+'\n</body>\n</html>\n')
print('wrote',os.path.normpath(out))
