"""Preprocess: no data injection needed (read-only databases only)."""
import argparse

def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("Инъекция данных не требуется — используются read-only источники данных")

if __name__ == "__main__":
    main()
