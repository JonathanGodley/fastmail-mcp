# Launcher: reads the working FASTMAIL_API_TOKEN from the local MCP client config
# and injects it into the child process environment. The value is never printed,
# logged, or written to disk. Usage: python scripts/probes/run-probe.py <probe.mjs>
import json, os, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))

cfg_path = os.path.expanduser('~/.claude.json')
with open(cfg_path, encoding='utf-8') as fh:
    cfg = json.load(fh)

token = cfg['mcpServers']['fastmail']['env']['FASTMAIL_API_TOKEN']

env = dict(os.environ)
env['FASTMAIL_API_TOKEN'] = token

script = sys.argv[1]
args = sys.argv[2:]
sys.exit(subprocess.run(['node', os.path.join(HERE, script)] + args, env=env, cwd=HERE).returncode)
