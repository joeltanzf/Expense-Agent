"""Build offline from standard CPython plus the pinned public dependencies.

Run with the Python installation to bundle. Do not use a venv as the base.
Output paths are explicit and a fresh directory is required.
"""
from pathlib import Path
import argparse, hashlib, importlib.metadata as md, json, os, shutil, subprocess, sys, zipfile
from packaging.requirements import Requirement

parser=argparse.ArgumentParser()
parser.add_argument('--out', required=True, type=Path)
args=parser.parse_args()
root=Path(__file__).resolve().parents[1]
out=args.out.resolve()
out.mkdir(parents=True,exist_ok=False)
payload=out/'payload'; payload.mkdir()
base=Path(sys.base_prefix)
runtime=payload/'runtime'; runtime.mkdir()
skip=shutil.ignore_patterns('__pycache__','*.pyc','site-packages','test','tests','idlelib','turtledemo','ensurepip','venv','sitecustomize.py','usercustomize.py')
shutil.copytree(base/'Lib',runtime/'Lib',ignore=skip)
shutil.copytree(base/'DLLs',runtime/'DLLs',ignore=shutil.ignore_patterns('__pycache__','*.pdb','*.lib'))
for name in ['python.exe','pythonw.exe','python3.dll','python312.dll','vcruntime140.dll','vcruntime140_1.dll','LICENSE.txt']:
    shutil.copy2(base/name,runtime/name)
site=runtime/'Lib/site-packages';site.mkdir()
seen={}
def copy_distribution(name):
    dist=md.distribution(name)
    key=dist.metadata['Name'].lower().replace('_','-')
    if key in seen:return
    seen[key]=dist.version
    for req in dist.requires or []:
        req=Requirement(req)
        if not req.marker or req.marker.evaluate({'extra':''}):copy_distribution(req.name)
    for rel in dist.files or []:
        rel=Path(str(rel))
        if '..' in rel.parts or '__pycache__' in rel.parts or rel.suffix=='.pyc':continue
        original=Path(dist.locate_file(rel)); dest=site/rel
        if original.is_file():
            dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(original,dest)
for name in ['pdfplumber','openpyxl']:copy_distribution(name)
for folder in ['app','scripts','docs']:
    shutil.copytree(root/folder,payload/folder,ignore=shutil.ignore_patterns('__pycache__','*.pyc'))
for name in ['README.md','requirements.txt','THIRD-PARTY-NOTICES.md']:
    shutil.copy2(root/name,payload/name)
(payload/'components.json').write_text(json.dumps({'python':sys.version.split()[0],'packages':seen},indent=2))
compiler=Path(os.environ['SystemRoot'])/'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
def compile(source,target,extra=()):
    command=[str(compiler),'/nologo','/target:winexe','/platform:x64','/optimize+', '/r:System.Windows.Forms.dll', '/r:System.dll', '/out:'+str(target), *extra, str(source)]
    subprocess.run(command,check=True)
compile(root/'build/Launcher.cs',payload/'TNG Expense Agent.exe')
manifest={str(p.relative_to(payload)).replace('\\','/'):hashlib.sha256(p.read_bytes()).hexdigest() for p in payload.rglob('*') if p.is_file()}
(payload/'manifest.json').write_text(json.dumps(manifest,indent=2))
archive=out/'payload.zip'
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as z:
    for p in sorted(payload.rglob('*')):
        if p.is_file():z.write(p,str(p.relative_to(payload)).replace('\\','/'))
exe=out/'TNG-Expense-Agent-Setup.exe'
framework=compiler.parent
compile(root/'build/Installer.cs',exe,[
    '/r:'+str(framework/'System.IO.Compression.dll'),'/r:'+str(framework/'System.IO.Compression.FileSystem.dll'),
    '/resource:'+str(archive)+',payload.zip'])
hash=hashlib.sha256(exe.read_bytes()).hexdigest()
(out/'SHA256SUMS.txt').write_text(hash+'  '+exe.name+'\n')
print(json.dumps({'installer':str(exe),'size_mb':round(exe.stat().st_size/1024**2,1),'packages':seen,'sha256':hash},indent=2))
