import argparse

def parse_args():
    parser = argparse.ArgumentParser(description="Convert CSV to formatted Excel")
    parser.add_argument("--input", required=True, help="Path to input CSV")
    parser.add_argument("--output", required=True, help="Path to output XLSX")
    return parser.parse_args()

if __name__ == "__main__":
    args = parse_args()
    print(f"Converting {args.input} to {args.output}")
