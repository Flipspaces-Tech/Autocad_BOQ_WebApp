# run_all.py
import sys, subprocess, datetime, pathlib

PY = sys.executable  # uses your current venv/python
ROOT = pathlib.Path(__file__).parent
LOGS = ROOT / "logs"
LOGS.mkdir(exist_ok=True)

SCRIPTS = [
    "one.py",
    # "measure_universal_envelope.py",
    # "check_dxf_metrics.py",
    "tr_2.py",
    "allinone.py",
    # "meger.py",
]

STOP_ON_ERROR = True  # set False to continue even if one fails

results = []
for script in SCRIPTS:
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log = LOGS / f"{pathlib.Path(script).stem}_{ts}.log"
    print(f"\n=== Running {script} ===")
    with log.open("w", encoding="utf-8") as f:
        proc = subprocess.run([PY, str(ROOT / script)],
                              stdout=subprocess.PIPE,
                              stderr=subprocess.STDOUT,
                              text=True)
        f.write(proc.stdout)
    print(f"→ exit {proc.returncode} | log: {log.name}")
    results.append((script, proc.returncode, log))
    if STOP_ON_ERROR and proc.returncode != 0:
        break

print("\nSummary:")
for s, code, log in results:
    print(f"  {s:<32} exit {code}  (log: {log.name})")
sys.exit(max((code for _, code, _ in results), default=0))
