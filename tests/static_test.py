from pathlib import Path
from html.parser import HTMLParser
import re,json
root=Path(__file__).resolve().parents[1]/'dist';text=(root/'index.html').read_text()
class Check(HTMLParser):
 def __init__(self):super().__init__();self.ids=[];self.refs=[];self.anchors=[]
 def handle_starttag(self,tag,attrs):
  a=dict(attrs)
  if a.get('id'):self.ids.append(a['id'])
  if tag in ('script','link'):
   r=a.get('src') or a.get('href');
   if r and not r.startswith(('http','//')):self.refs.append(r)
c=Check();c.feed(text)
assert len(c.ids)==len(set(c.ids)),set(x for x in c.ids if c.ids.count(x)>1)
for r in c.refs:assert (root/r).exists(),r
js=(root/'assets'/'app.js').read_text()
for id in re.findall(r"\$\('([^']+)'\)",js):assert id in c.ids, 'missing DOM ID '+id
assert '@import' not in (root/'assets'/'styles.css').read_text()
assert 'clientId: \'\'' in (root/'config.js').read_text()
assert not any(x in (root/'config.js').read_text().lower() for x in ['client_secret','password=','bearer '])
print('PASS static — unique IDs, resolved asset refs, DOM dependencies, no external fonts, no credentials')
