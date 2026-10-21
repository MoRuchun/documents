#!/usr/bin/env python3
"""Install or update documents, back up local copies and disable bundled duplicates."""
from __future__ import annotations
import argparse
from datetime import datetime
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from urllib.request import urlopen
import venv
from zipfile import ZipFile

REPOSITORY='https://github.com/MoRuchun/documents'
IGNORED=shutil.ignore_patterns('.git','__pycache__','*.pyc','.DS_Store','.venv','install-state.json')


def is_documents(directory):
    file=directory/'SKILL.md'
    if not file.is_file():return False
    text=file.read_text(encoding='utf-8')
    header=text.split('---',2)[1] if text.startswith('---') else ''
    return re.search(r'''(?m)^name:\s*["']?documents["']?\s*$''',header) is not None


def canonical(path):
    return os.path.normcase(str(Path(path).expanduser().resolve()))


def parse_path(text):
    match=re.search(r'''(?m)^\s*path\s*=\s*("(?:\\.|[^"\\])*"|'[^']*')''',text)
    if not match:return None
    value=match.group(1)
    return json.loads(value) if value.startswith('"') else value[1:-1]


def update_config(text, overrides):
    """Edit matching skill entries only; retain other configuration and comments."""
    targets={canonical(p):(str(p),enabled) for p,enabled in overrides.items()}
    section=re.compile(r'(?m)^[ \t]*\[\[skills\.config\]\][ \t]*(?:#.*)?$')
    changes=[];seen=set()
    for match in section.finditer(text):
        next_header=re.search(r'(?m)^[ \t]*\[',text[match.end():])
        end=match.end()+next_header.start() if next_header else len(text)
        block=text[match.start():end]
        value=parse_path(block)
        if value is None:continue
        path=Path(value).expanduser()
        normalized=canonical(path/'SKILL.md' if path.name!='SKILL.md' else path)
        if normalized not in targets:continue
        seen.add(normalized)
        enabled='true' if targets[normalized][1] else 'false'
        if re.search(r'(?m)^[ \t]*enabled\s*=',block):
            block=re.sub(r'(?m)^([ \t]*enabled\s*=\s*)(?:true|false)([^\n]*)$',lambda m:m.group(1)+enabled+m.group(2),block)
        else:
            block=block.rstrip()+'\nenabled = '+enabled+'\n\n'
        changes.append((match.start(),end,block))
    result=text
    for start,end,block in reversed(changes):result=result[:start]+block+result[end:]
    for key,(path,enabled) in targets.items():
        if key in seen:continue
        result=result.rstrip()+'\n\n[[skills.config]]\npath = '+json.dumps(path,ensure_ascii=False)+'\nenabled = '+str(enabled).lower()+'\n'
    if sys.version_info>=(3,11):
        import tomllib
        tomllib.loads(result)
    return result


def write_atomic(file,text):
    file.parent.mkdir(parents=True,exist_ok=True)
    mode=file.stat().st_mode if file.exists() else 0o600
    fd,name=tempfile.mkstemp(prefix=file.name+'.',dir=file.parent)
    tmp=Path(name)
    try:
        with os.fdopen(fd,'w',encoding='utf-8') as stream:stream.write(text)
        tmp.chmod(mode)
        tmp.replace(file)
    finally:
        if tmp.exists():tmp.unlink()


def discover(home,codex_home,destination):
    locals_found=[]
    for base in [home/'.agents/skills',codex_home/'skills']:
        if not base.is_dir():continue
        for directory in base.iterdir():
            if directory.is_dir() and is_documents(directory):locals_found.append(directory.absolute())
    if destination.exists() or destination.is_symlink():
        if not is_documents(destination):
            raise ValueError(f'目标目录已经存在且不是 documents 技能：{destination}')
        locals_found.append(destination.absolute())
    plugins=[]
    cache=codex_home/'plugins/cache'
    if cache.is_dir():
        for file in cache.glob('**/skills/*/SKILL.md'):
            if is_documents(file.parent):plugins.append(file.absolute())
    return sorted(set(locals_found)),sorted(set(plugins))


def prepare_python(codex_home,requirements,skip):
    if skip or all(importlib.util.find_spec(name) is not None for name in ['docx','lxml']):
        return sys.executable
    folder=codex_home/'skill-envs/documents'
    executable=folder/('Scripts/python.exe' if os.name=='nt' else 'bin/python')
    if not executable.exists():venv.EnvBuilder(with_pip=True).create(folder)
    subprocess.run([str(executable),'-m','pip','install','-r',str(requirements)],check=True)
    return str(executable)


def install(source,home,codex_home,destination=None,skip_deps=False,dry_run=False):
    source=source.expanduser().resolve();home=home.expanduser().resolve();codex_home=codex_home.expanduser().resolve()
    destination=(destination or home/'.agents/skills/documents').expanduser().absolute()
    if not is_documents(source):raise ValueError(f'源目录缺少 documents 的 SKILL.md：{source}')
    if not (source/'scripts/user_formatting.py').is_file():raise ValueError('源目录缺少格式工具')
    if source!=destination.resolve() and source in destination.resolve().parents:
        raise ValueError('安装目标应放在源目录之外')
    existing,plugins=discover(home,codex_home,destination)
    config=codex_home/'config.toml'
    original=config.read_text(encoding='utf-8') if config.exists() else None
    overrides={p:False for p in plugins}
    overrides[destination/'SKILL.md']=True
    updated=update_config(original or '',overrides)
    plan={'repository':REPOSITORY,'destination':str(destination),
          'local_copies_to_back_up':[str(p) for p in existing],
          'bundled_skills_to_disable':[str(p) for p in plugins],
          'config':str(config),'restart_codex':True}
    if dry_run:return dict(plan,dry_run=True)
    # Stage before moving anything, including when this script runs from destination.
    with tempfile.TemporaryDirectory(prefix='documents-install-') as temporary:
        staged=Path(temporary)/'documents'
        shutil.copytree(source,staged,ignore=IGNORED)
        python=prepare_python(codex_home,staged/'requirements.txt',skip_deps)
        backup_base=codex_home/'skill-backups'
        backup_base.mkdir(parents=True,exist_ok=True)
        backup=Path(tempfile.mkdtemp(prefix='documents-'+datetime.now().strftime('%Y%m%d-%H%M%S')+'-',dir=backup_base))
        if original is not None:shutil.copy2(config,backup/'config.toml')
        moved=[];created=False
        try:
            for i,old in enumerate(existing):
                saved=backup/f'local-{i+1}'
                shutil.move(str(old),str(saved));moved.append((old,saved))
            destination.parent.mkdir(parents=True,exist_ok=True)
            created=True
            shutil.copytree(staged,destination)
            write_atomic(config,updated)
            state=dict(plan,backup_directory=str(backup),python_executable=python,
                       backups=[{'original':str(a),'backup':str(b)} for a,b in moved])
            (destination/'install-state.json').write_text(json.dumps(state,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
            (backup/'install-state.json').write_text(json.dumps(state,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
            return state
        except BaseException:
            if created and destination.exists():shutil.rmtree(destination)
            for old,saved in reversed(moved):
                if saved.exists() or saved.is_symlink():shutil.move(str(saved),str(old))
            if original is None:
                if config.exists():config.unlink()
            else:write_atomic(config,original)
            raise


def download_latest(ref,temporary):
    url=f'https://codeload.github.com/MoRuchun/documents/zip/{ref}'
    archive=temporary/'source.zip'
    with urlopen(url,timeout=30) as response, archive.open('wb') as output:
        shutil.copyfileobj(response,output)
    extracted=temporary/'source';extracted.mkdir()
    with ZipFile(archive) as z:
        for item in z.infolist():
            target=(extracted/item.filename).resolve()
            if extracted.resolve() not in target.parents:raise ValueError('下载包包含目录之外的路径')
        z.extractall(extracted)
    roots=[p for p in extracted.iterdir() if p.is_dir() and is_documents(p)]
    if len(roots)!=1:raise ValueError('下载包中没有唯一的 documents 技能目录')
    return roots[0]


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--source',type=Path,default=Path(__file__).resolve().parents[1])
    ap.add_argument('--home',type=Path,default=Path.home(),help='用户目录；用于自定义安装位置或隔离检查')
    ap.add_argument('--codex-home',type=Path,default=None)
    ap.add_argument('--dest',type=Path,default=None)
    ap.add_argument('--update',action='store_true',help='从 GitHub 下载最新版本并替换安装')
    ap.add_argument('--ref',default='main')
    ap.add_argument('--skip-deps',action='store_true',help='跳过 Python 依赖准备')
    ap.add_argument('--dry-run',action='store_true',help='仅显示计划，不写入文件或安装依赖')
    args=ap.parse_args()
    ch=args.codex_home or Path(os.environ.get('CODEX_HOME',str(args.home/'.codex')))
    try:
        if args.update:
            with tempfile.TemporaryDirectory(prefix='documents-download-') as td:
                source=download_latest(args.ref,Path(td))
                result=install(source,args.home,ch,args.dest,args.skip_deps,args.dry_run)
        else:result=install(args.source,args.home,ch,args.dest,args.skip_deps,args.dry_run)
    except (OSError,ValueError,subprocess.CalledProcessError) as exc:
        ap.exit(1,f'安装失败：{exc}\n')
    print(json.dumps(result,ensure_ascii=False,indent=2))
    if not args.dry_run:print('安装完成。请重启 Codex 后使用 $documents。')

if __name__=='__main__':main()
