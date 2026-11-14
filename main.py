import argparse
import pandas as pd

def parse_args():
    parser = argparse.ArgumentParser(description="Convert CSV to formatted Excel")
    parser.add_argument("--input", required=True, help="Path to input CSV")
    parser.add_argument("--output", required=True, help="Path to output XLSX")
    return parser.parse_args()

def main():
    args = parse_args()
    print(f"Reading {args.input}...")
    try:
        df = pd.read_csv(args.input)
    except FileNotFoundError:
        print(f"Error: File {args.input} not found.")
        return
    print(f"Loaded {len(df)} rows.")

    print(f"Writing to {args.output}...")
    with pd.ExcelWriter(args.output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        worksheet = writer.sheets['Data']
        
        # Автоширина колонок
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2
            
    print("Done!")

if __name__ == "__main__":
    main()
