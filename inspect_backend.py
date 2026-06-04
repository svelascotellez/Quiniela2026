import openpyxl

def main():
    wb_path = r"C:\Users\salva\Downloads\Quiniela_Mundial_2026_Final_vacia (1).xlsx"
    wb = openpyxl.load_workbook(wb_path, data_only=False)
    ws_backend = wb['Backend']
    
    print("--- 'Backend' formulas ---")
    for r in range(1, 30):
        row_vals = [ws_backend.cell(row=r, column=c).value for c in range(1, 6)]
        print(f"Row {r}: {row_vals}")

if __name__ == "__main__":
    main()
