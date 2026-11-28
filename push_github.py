"""Push files from the local skill dir to FrancesNK/workbuddy-skills (macOS skill).

Token lives OUTSIDE the skill dir at ~/.workbuddy/.github_pat (never commit it).

Usage:
  python3 push_github.py                       # push SKILL.md
  python3 push_github.py rel/path another.py   # push specific files (paths relative to skill dir)
"""
import base64
import json
import os
import sys
import urllib.request

REPO = 'FrancesNK/workbuddy-skills'
BRANCH = 'main'
SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.expanduser('~/.workbuddy/.github_pat')


def api(url, data=None, method=None):
    req = urllib.request.Request(url, headers={
        'Authorization': f'Bearer {open(TOKEN_FILE).read().strip()}',
        'User-Agent': 'curl',
    }, data=data, method=method)
    return json.load(urllib.request.urlopen(req, timeout=60))


def push(rel):
    local = os.path.join(SKILL_DIR, rel)
    url = f'https://api.github.com/repos/{REPO}/contents/{rel}'
    try:
        sha = api(url)['sha']
    except Exception:
        sha = None
    body = {
        'message': f'Update {rel} from local skill',
        'content': base64.b64encode(open(local, 'rb').read()).decode(),
        'branch': BRANCH,
    }
    if sha:
        body['sha'] = sha
    r = api(url, json.dumps(body).encode(), 'PUT')
    print('PUSHED', rel, '->', r['commit']['sha'][:8])


if __name__ == '__main__':
    files = sys.argv[1:] or ['SKILL.md']
    for f in files:
        push(f)
