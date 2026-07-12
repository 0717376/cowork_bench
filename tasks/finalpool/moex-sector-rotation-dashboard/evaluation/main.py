from argparse import ArgumentParser

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from check_local import run_checks

ACCURACY_THRESHOLD = 70.0

if __name__ == "__main__":
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--res_log_file", required=False, help="Path to result log file")
    parser.add_argument("--launch_time", required=False, help="Launch time")
    args = parser.parse_args()

    try:
        results, fatal = run_checks(args.agent_workspace, args.groundtruth_workspace)
    except Exception as e:
        print("local check error: ", e)
        sys.exit(1)

    if fatal:
        print("local check failed: ", fatal)
        sys.exit(1)

    if not results:
        print("local check failed: no checks produced")
        sys.exit(1)

    # --- CRITICAL gate: any critical failure => immediate FAIL ---
    critical = [r for r in results if r["critical"]]
    critical_failed = [r for r in critical if not r["passed"]]
    if critical_failed:
        print("CRITICAL checks FAILED:")
        for r in critical_failed:
            print(f"  - {r['name']}: {r['msg']}")
        sys.exit(1)

    # --- Accuracy gate over all checks ---
    total = len(results)
    passed = sum(1 for r in results if r["passed"])
    accuracy = passed / total * 100.0

    print(f"Critical checks: {len(critical)} (all passed)")
    print(f"Accuracy: {passed}/{total} = {accuracy:.1f}%")
    for r in results:
        if not r["passed"]:
            print(f"  FAIL {r['name']}: {r['msg']}")

    if accuracy < ACCURACY_THRESHOLD:
        print(f"Accuracy {accuracy:.1f}% below threshold {ACCURACY_THRESHOLD}%")
        sys.exit(1)

    print("Pass all tests!")
