"""
Preprocess for q4-sales-reconciliation task.
ClickHouse (sf_data schema) is read-only and globally seeded/russified, so no
data injection or clearing is needed.
"""
import argparse


def main():
    # No writable schemas to DELETE - read-only data source (ClickHouse sf_data).
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    print("No data injection needed - ClickHouse (sf_data) is read-only")


if __name__ == "__main__":
    main()
