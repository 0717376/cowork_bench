import argparse


def main():
    # Нет записываемых схем для очистки — moex.* read-only, глобально засеяна.
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("No preprocess needed for read-only MOEX Finance task")


if __name__ == "__main__":
    main()
