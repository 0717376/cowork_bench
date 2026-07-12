"""Preprocess: no data injection needed (read-only ClickHouse HR_ANALYTICS only)."""
import argparse

def main():
    # No writable schemas to DELETE - read-only data sources (ClickHouse HR_ANALYTICS).
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("No data injection needed - using read-only data")

if __name__ == "__main__":
    main()
