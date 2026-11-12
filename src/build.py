import json
names=['Book','Manage','Admin','Pay','Privacy','Styles','Common']
code=open('Code.gs').read()
pages='\n\n/* ============================== Pages (bundled) ============================== */\nconst PAGES = {\n' + ',\n'.join(f'  {n}: {json.dumps(open(n+".html").read())}' for n in names) + '\n};\n'
open('../dist/Code.gs','w').write(code+pages)
print(len(code+pages))
