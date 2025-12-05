import pandas as pd
import os
import sys

# Configurar la codificación de salida a UTF-8
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def leer_hojas_especificas():
    # Ruta del archivo Excel
    ruta_excel = os.path.join(os.path.dirname(__file__), 'CAJA DE AHORRO COMUNAL VERDECOCHA ORIGINAL MDF.xlsx')
    
    try:
        # Leer el archivo Excel
        print(f"Leyendo archivo: {ruta_excel}")
        xls = pd.ExcelFile(ruta_excel)
        
        # Lista de hojas a leer
        hojas_a_leer = [
            'LISTADO DE SOCIOS',
            'REGISTRO CONTABLE',
            'RAPA',
            'CAJA CHICA',
            'INTERÉS',
            'CÁLCULO INTERES',
            'SUMA INTERES TOTAL POR SOCIO',
            'LIBRETAS'
        ]
        
        # Verificar qué hojas existen en el archivo
        hojas_disponibles = xls.sheet_names
        print(f"\nHojas disponibles en el archivo: {hojas_disponibles}")
        
        # Filtrar solo las hojas que existen
        hojas_a_leer = [h for h in hojas_a_leer if h in hojas_disponibles]
        
        if not hojas_a_leer:
            print("No se encontraron las hojas especificadas en el archivo.")
            return
            
        print(f"\nLeyendo las siguientes hojas: {hojas_a_leer}")
        
        # Leer cada hoja específica
        for hoja in hojas_a_leer:
            print(f"\n{'='*80}")
            print(f"ANÁLISIS DE LA HOJA: {hoja}")
            print(f"{'='*80}")
            
            try:
                # Leer la hoja con manejo de codificación
                try:
                    df = pd.read_excel(xls, sheet_name=hoja, header=None, engine='openpyxl')
                except Exception as e:
                    print(f"Error al leer la hoja con openpyxl: {e}")
                    # Intentar con otro motor si falla
                    df = pd.read_excel(xls, sheet_name=hoja, header=None, engine='xlrd')
                
                # Mostrar información básica
                print(f"\n📊 Dimensiones: {df.shape[0]} filas × {df.shape[1]} columnas")
                
                # Mostrar las primeras filas con datos
                print("\n📝 Primeras filas con datos:")
                # Usar un formato alternativo para evitar problemas de codificación
                for i, row in df.head(10).iterrows():
                    print(f"Fila {i}:")
                    for col in df.columns:
                        if pd.notna(row[col]):
                            print(f"  Col {col}: {row[col]}")
                
                # Contar valores no nulos por columna
                print("\n🔍 Valores no nulos por columna:")
                print(df.count())
                
                # Identificar posibles columnas con datos importantes
                print("\n🔍 Columnas con más del 50% de datos no nulos:")
                columnas_importantes = df.columns[df.count() > (len(df) * 0.5)].tolist()
                print(columnas_importantes)
                
                # Si hay columnas importantes, mostrarlas
                if columnas_importantes:
                    print("\n📋 Vista previa de columnas importantes:")
                    print(df[columnas_importantes].head())
                
            except Exception as e:
                print(f"❌ Error al procesar la hoja {hoja}: {str(e)}")
    
    except Exception as e:
        print(f"Error al leer el archivo: {str(e)}")

if __name__ == "__main__":
    leer_hojas_especificas()
