import pandas as pd
import win32com.client as win32
import os
import time

RUTA_EXCEL = r'C:\Users\PC de Johan\Downloads\certificados\prueba1.xlsx'
MODO_PRUEBA = True   # True = abre borradores sin enviar

# 1. CONECTAR CON OUTLOOK
outlook = win32.Dispatch('Outlook.Application')
print("Outlook conectado")

# 2. LEER EXCEL
df = pd.read_excel(RUTA_EXCEL, dtype={'Cedula': str})
df['Correo'] = df['Correo'].str.strip()
df['ArchivoPDF'] = df['ArchivoPDF'].str.strip()
print(f"Registros leídos: {len(df)}")

# 3. VERIFICAR QUE LOS PDF EXISTAN (antes de enviar nada)
faltantes = df[~df['ArchivoPDF'].apply(os.path.isfile)]
if len(faltantes):
    print(f"\nATENCIÓN: {len(faltantes)} PDF(s) no encontrados:")
    for _, f in faltantes.iterrows():
        print(f"   {f['Nombre']}: {f['ArchivoPDF']}")
    print("\nCorrija las rutas antes de continuar.")
    exit()
print("Todos los PDF existen\n")

# 4. RECORRER CADA REGISTRO Y ENVIAR
for i, fila in df.iterrows():
    nombre  = fila['Nombre']
    correo  = fila['Correo']
    carrera = fila['Carrera']
    pdf     = fila['ArchivoPDF']

    try:
        mail = outlook.CreateItem(0)
        mail.To      = correo
        mail.Subject = f'Constancias - Egresados Embajadores - {nombre}'
        mail.Body    = f"""Buenas tardes, {nombre}.

Desde el programa de Egresados de la Escuela Tecnológica Instituto Técnico Central (ETITC), nos complace hacer entrega de su constancia como Egresado Embajador, en reconocimiento a su compromiso. 
Reciba nuestras más sinceras felicitaciones y agradecimientos por representar con orgullo los valores y principios de nuestra institución.

Cordialmente,
Jhon Fredy Ortegate
Profesional del Programa de Egresados
Escuela Tecnológica Instituto Técnico Central (ETITC)"""

        mail.Attachments.Add(pdf)

        if MODO_PRUEBA:
            mail.Display()
            print(f"[{i+1}/{len(df)}] BORRADOR: {nombre} -> {correo}")
        else:
            mail.Send()
            print(f"[{i+1}/{len(df)}] ENVIADO: {nombre} -> {correo}")
            time.sleep(3)

    except Exception as e:
        print(f"[{i+1}/{len(df)}] ERROR: {nombre} -> {correo}: {e}")

print("\nProceso terminado")