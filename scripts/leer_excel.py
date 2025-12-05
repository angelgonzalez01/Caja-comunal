import pandas as pd
import os

def leer_archivo_excel():
    # Ruta del archivo Excel
    ruta_excel = os.path.join(os.path.dirname(__file__), 'CAJA DE AHORRO COMUNAL VERDECOCHA ORIGINAL MDF.xlsx')
    
    try:
        # Leer el archivo Excel
        print(f"Leyendo archivo: {ruta_excel}")
        xls = pd.ExcelFile(ruta_excel)
        
        # Obtener nombres de las hojas
        print(f"\nHojas encontradas: {xls.sheet_names}")
        
        # Leer cada hoja
        for sheet_name in xls.sheet_names:
            print(f"\n--- Contenido de la hoja: {sheet_name} ---")
            df = pd.read_excel(xls, sheet_name=sheet_name)
            print(f"Filas: {len(df)}, Columnas: {len(df.columns) if len(df) > 0 else 0}")
            if not df.empty:
                print("\nPrimeras filas:")
                print(df.head())
                print("\nColumnas:", list(df.columns))
            else:
                print("La hoja está vacía.")
            print("-" * 50)
            
    except Exception as e:
        print(f"Error al leer el archivo: {str(e)}")

if __name__ == "__main__":
    leer_archivo_excel()
