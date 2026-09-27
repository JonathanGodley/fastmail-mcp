# Launcher: reads the working Fastmail credentials from the local MCP client config
# and injects them into the child process environment. Values are never printed,
# logged, or written to disk. Usage: python scripts/probes/run-probe.py <probe.mjs>
#
# FASTMAIL_API_TOKEN is required. The CalDAV credentials and display name are
# injected only when the config carries them (a JMAP-only setup has no calendar app
# password), and a missing one is left to the probe to report. FASTMAIL_TIMEZONE is
# optional too: when unset, the server silently falls back to the host's own zone.
import json, os, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))

cfg_path = os.path.expanduser('~/.claude.json')
with open(cfg_path, encoding='utf-8') as fh:
    cfg = json.load(fh)

cfg_env = cfg['mcpServers']['fastmail']['env']

env = dict(os.environ)
env['FASTMAIL_API_TOKEN'] = cfg_env['FASTMAIL_API_TOKEN']

for key in ('FASTMAIL_CALDAV_USERNAME', 'FASTMAIL_CALDAV_PASSWORD', 'FASTMAIL_CALDAV_DISPLAY_NAME', 'FASTMAIL_TIMEZONE'):
    value = cfg_env.get(key)
    if value:
        env[key] = value

script = sys.argv[1]
args = sys.argv[2:]
sys.exit(subprocess.run(['node', os.path.join(HERE, script)] + args, env=env, cwd=HERE).returncode)
