# Converts every .docx and .pptx under tests/ with one CLI build and records the PDF hash,
# wall time and peak memory, so a speed change can be checked for byte-identical output.
#
#   python3 perf/hash_fixtures.py <cli-binary> before.json
#   python3 perf/hash_fixtures.py <cli-binary> after.json --compare before.json
#
# Run it from a directory whose Fonts/ folder holds the fonts you deploy with, or pass --cwd.
import argparse, hashlib, json, os, subprocess, sys, tempfile, time
from concurrent.futures import ThreadPoolExecutor

parser = argparse.ArgumentParser()
parser.add_argument('cli')
parser.add_argument('out')
parser.add_argument('--compare')
parser.add_argument('--cwd', default=os.getcwd())
parser.add_argument('--threads', type=int, default=max(1, (os.cpu_count() or 2) // 2))
args = parser.parse_args()

root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
cli = os.path.abspath(args.cli)
files = sorted(os.path.join(d, f) for d, _, fs in os.walk(os.path.join(root, 'tests')) for f in fs if f.endswith(('.docx', '.pptx')))
# ru_maxrss is in bytes on macOS and in kilobytes on Linux.
rss_unit = 1 if sys.platform == 'darwin' else 1024


def convert(path):
    name = os.path.relpath(path, root)
    with tempfile.TemporaryDirectory() as t:
        out = os.path.join(t, 'out.pdf')
        start = time.perf_counter()
        p = subprocess.Popen([cli, 'convert', path, out], cwd=args.cwd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        _, status, usage = os.wait4(p.pid, 0)
        ms = int((time.perf_counter() - start) * 1000)
        sha = hashlib.sha256(open(out, 'rb').read()).hexdigest() if os.path.exists(out) else None
        return name, {'exit': os.waitstatus_to_exitcode(status), 'ms': ms, 'rss_mib': usage.ru_maxrss * rss_unit >> 20, 'sha': sha}


with ThreadPoolExecutor(args.threads) as pool:
    results = dict(pool.map(convert, files))
json.dump(results, open(args.out, 'w'), indent=0)
print(f"{len(results)} files, {sum(r['ms'] for r in results.values()) / 1000:.0f} s total")

if args.compare:
    before = json.load(open(args.compare))
    changed = sorted(k for k in results if k in before and results[k]['sha'] != before[k]['sha'])
    for k in changed:
        print('changed', k)
    print(f'{len(changed)} of {len(results)} outputs changed')
    sys.exit(1 if changed else 0)
