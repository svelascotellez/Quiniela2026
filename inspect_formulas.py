import openpyxl

def main():
    wb_path = r"C:\Users\salva\Downloads\Quiniela_Mundial_2026_Final_vacia (1).xlsx"
    wb = openpyxl.load_workbook(wb_path, data_only=False)
    ws_ko = wb['Eliminatorias']
    
    # Check row 6 Columns C, D, E, F, G, H, I, J
    print("Row 6 cells:")
    for c in range(1, 11):
        cell = ws_ko.cell(row=6, column=c)
        print(f"Col {c}: {cell.value}")

if __name__ == "__main__":
    main()
