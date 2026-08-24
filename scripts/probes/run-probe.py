# Launcher: reads the working Fastmail credentials from the local MCP client config
# and injects them into the child process environment. Values are never printed,
# logged, or written to disk. Usage: python scripts/probes/run-probe.py <probe.mjs>
#
# FASTMAIL_API_TOKEN is required (the JMAP probes cannot run without it). The CalDAV
# username/password and the CalDAV display name are optional and only injected when
# the config carries them, because the calendar probes need a separate app password
# that a JMAP-only setup will not have. A missing CalDAV credential is left to the
# probe to report, so a JMAP probe still runs on a config that has no calendar
# access configured.
#
# FASTMAIL_TIMEZONE is forwarded the same way but is not calendar-scoped - it also
# governs email `date` rendering (see README.md), so a JMAP-only setup routinely
# sets it. It is optional here because the server itself falls back to the host's
# own zone when it is unset; unlike a missing CalDAV credential, that fallback is
# silent, not something a probe reports.
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
