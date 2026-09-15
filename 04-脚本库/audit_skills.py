# -*- coding: utf-8 -*-
"""静态安全扫描：第三方技能包"""
import os, re, json, sys, collections

ROOT = sys.argv[1] if len(sys.argv) > 1 else r'D:\<化工运行时>\pkg'
PATTERNS = {
    'shell_exec': r'\b(os\.system|os\.popen|subprocess\.(run|Popen|call|check_output|check_call))\b',
    'dynamic_exec': r'(?<![\w.])(eval|exec)\s*\(',
    'dynamic_import': r'__import__\s*\(',
    'network': r'\b(requests\.(get|post|Session)|urllib\.request|http\.client|socket\.socket|ftplib|aiohttp|httpx|websocket)\b',
    'winreg': r'\bwinreg\b',
    'ctypes': r'\bctypes\b',
    'destructive': r'\b(shutil\.rmtree|os\.remove|os\.unlink|os\.rmdir|os\.removedirs)\b',
    'b64_decode': r'base64\.b64decode',
    'pickle_load': r'pickle\.loads?\s*\(',
}
CRED = re.compile(r'(?i)(api[_-]?key|access[_-]?token|secret[_-]?key|password\s*=|passwd\s*=|private[_-]?key)')
INJECT = re.compile(r'(?i)(ignore (all )?(previous|prior|above) instructions|忽略(之前|以上|前面)的?(指令|要求)|不需要?用户确认|without user (confirmation|consent)|auto[- ]?execute without)')
BIN_EXT = ('.exe', '.dll', '.bat', '.cmd', '.ps1', '.vbs', '.jar', '.scr', '.msi', '.com')
TEXT_EXT = ('.py', '.json', '.md', '.txt', '.yaml', '.yml', '.toml', '.cfg', '.ini', '.csv')

hits = collections.defaultdict(list)
urls = collections.Counter()
binary, skills, scripts = [], [], []
inject = []
cred = []
n_files = 0

for dirpath, dirnames, filenames in os.walk(ROOT):
    for fn in filenames:
        p = os.path.join(dirpath, fn)
        rel = os.path.relpath(p, ROOT)
        n_files += 1
        low = fn.lower()
        if low.endswith(BIN_EXT):
            binary.append(rel)
        if fn == 'SKILL.md':
            skills.append(rel)
        if low.endswith('.py'):
            scripts.append(rel)
        if low.endswith(TEXT_EXT):
            try:
                t = open(p, encoding='utf-8', errors='ignore').read()
            except OSError:
                continue
            for name, pat in PATTERNS.items():
                for m in re.finditer(pat, t):
                    if len(hits[name]) < 400:
                        hits[name].append((rel, m.group(0)))
            for m in re.finditer(r'https?://[^\s"\'\)\]<>]+', t):
                urls[m.group(0)[:90]] += 1
            if fn.endswith('.md'):
                for m in INJECT.finditer(t):
                    inject.append((rel, m.group(0)))
            if low.endswith(('.py', '.json', '.yaml', '.yml', '.cfg', '.ini')):
                for m in CRED.finditer(t):
                    cred.append((rel, m.group(0)))

out = {
    'root': ROOT, 'total_files': n_files,
    'skill_md_count': len(skills),
    'python_scripts': len(scripts),
    'binary_scripts': binary[:40],
    'binary_count': len(binary),
    'pattern_hits': {k: {'count': len(v), 'samples': v[:6]} for k, v in hits.items()},
    'prompt_injection_suspects': inject[:20],
    'credential_patterns': cred[:20],
    'external_urls_top': urls.most_common(25),
    'url_count': len(urls),
}
print(json.dumps(out, ensure_ascii=False, indent=2))
