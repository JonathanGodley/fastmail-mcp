# Launcher: reads the working Fastmail credentials and injects them into the child
# process environment, because the harness gives the spawned server an empty home
# where it would not find them itself. Values are never printed, logged, or written
# to disk. Usage: python scripts/probes/run-probe.py <probe.mjs>
#
# Each variable comes from ~/.fastmail-mcp/.env, else the fastmail entry in
# ~/.claude.json; either may be absent.
#
# FASTMAIL_API_TOKEN is required. The CalDAV credentials and display name are
# injected only when a source carries them (a JMAP-only setup has no calendar app
# password), and a missing one is left to the probe to report. FASTMAIL_TIMEZONE is
# optional too: when unset, the server silently falls back to the host's own zone.
import json, os, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
KEYS = ('FASTMAIL_API_TOKEN', 'FASTMAIL_CALDAV_USERNAME', 'FASTMAIL_CALDAV_PASSWORD',
        'FASTMAIL_CALDAV_DISPLAY_NAME', 'FASTMAIL_TIMEZONE')
ENV_PATH = os.path.expanduser('~/.fastmail-mcp/.env')
CFG_PATH = os.path.expanduser('~/.claude.json')


def read_env_file(path):
    try:
        with open(path, encoding='utf-8-sig') as fh:
            lines = fh.read().splitlines()
    except FileNotFoundError:
        return {}
    out = {}
    for line in lines:
        line = line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        if line.startswith('export '):
            line = line[len('export '):]
        key, value = line.split('=', 1)
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in '"\'':
            value = value[1:-1]
        out[key.strip()] = value
    return out


def read_client_config(path):
    try:
        with open(path, encoding='utf-8') as fh:
            cfg = json.load(fh)
    except FileNotFoundError:
        return {}
    return cfg.get('mcpServers', {}).get('fastmail', {}).get('env', {})


file_env = read_env_file(ENV_PATH)
cfg_env = read_client_config(CFG_PATH)

env = dict(os.environ)
for key in KEYS:
    value = file_env.get(key) or cfg_env.get(key)
    if value:
        env[key] = value

if not (file_env.get('FASTMAIL_API_TOKEN') or cfg_env.get('FASTMAIL_API_TOKEN')):
    sys.exit(f'FASTMAIL_API_TOKEN not found in {ENV_PATH} or the fastmail entry in {CFG_PATH}')

script = sys.argv[1]
args = sys.argv[2:]
sys.exit(subprocess.run(['node', os.path.join(HERE, script)] + args, env=env, cwd=HERE).returncode)
