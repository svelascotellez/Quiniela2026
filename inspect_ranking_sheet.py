import openpyxl

def main():
    wb_path = r"C:\Users\salva\Downloads\Quiniela_Mundial_2026_Final_vacia (1).xlsx"
    wb = openpyxl.load_workbook(wb_path, data_only=True)
    ws_rf = wb['Ranking FIFA']
    
    # Check rows
    populated_rows = 0
    empty_rows = 0
    
    # We can write output to a file in utf-8 to avoid console encoding crashes
    with open("C:/Users/salva/.antigravity/Quiniela2026/ranking_sheet_info.txt", "w", encoding="utf-8") as f:
        f.write(f"Ranking FIFA Dimensions: {ws_rf.dimensions}\n")
        for r in range(1, 215):
            row_vals = [ws_rf.cell(row=r, column=c).value for c in range(1, 9)]
            if any(val is not None for val in row_vals):
                populated_rows += 1
                if r <= 10:
                    f.write(f"Row {r}: {row_vals}\n")
            else:
                empty_rows += 1
        f.write(f"Populated rows: {populated_rows}\n")
        f.write(f"Empty rows: {empty_rows}\n")
        
    print(f"Inspection complete. Populated: {populated_rows}, Empty: {empty_rows}")

if __name__ == "__main__":
    main()
