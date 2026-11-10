# INEScan 2x

Escanea frente y reverso de una INE (o cualquier credencial) con un escáner de Windows (WIA, probado con HP Smart Tank 525) y genera un PDF listo para imprimir:

- **PDF al doble (2x):** la credencial al doble de su tamaño real, una debajo de la otra.
- **PDF tamaño normal:** frente y reverso lado a lado, a 3 cm del borde superior.

Recorta y endereza solo, con margen blanco uniforme. El tamaño se toma de las medidas reales del escaneo.

## Documento oficio en 2 partes
Para documentos más largos que el cristal del escáner o que la hoja de tu impresora (p. ej. un acta en tamaño oficio). Botón **Documento oficio…**:

1. Escanea la parte de **arriba**.
2. Desliza el documento y escanea lo que falta, dejando ~2 cm repetidos de lo ya escaneado.
3. La app detecta dónde continúa el documento (puntos de referencia), alinea las dos partes, empareja el brillo y las une.
4. Detecta lo que realmente es papel y recorta el cristal vacío; endereza la inclinación y deja los bordes rectos. Con papel de color usa el color del papel; con papel blanco recorta al contenido (margen de 10 mm).
5. Elige la forma de impresión y genera el PDF:
   - **Oficio, tamaño real** (21.6 x 34 cm)
   - **Ajustar a 1 hoja carta/A4** (reducido)
   - **Dividir en hojas carta/A4** (tamaño real, cortando por una fila lo más vacía posible)

## Compatibilidad de escáneres
- **Escáneres que Windows reconoce (WIA):** USB o red, de cualquier marca cuyo driver use WIA. La app lee lo que ofrece cada driver (resolución, formatos, cristal plano) y se adapta. Probado con HP Smart Tank 520/525.
- **Impresoras de red que Windows no lista:** botón **Red…**, escribe la IP de la impresora y escanea por eSCL (AirScan/Mopria), sin drivers. Probado solo contra un servidor simulado.
- **Sin escáner o solo TWAIN:** usa **Abrir imagen…** con un escaneo guardado.
- Solo Windows.

## Uso
```
pip install -r requirements.txt
python escaneo_ine.py
```
o doble clic en `Abrir_Escaneo.bat`. Los PDF se guardan en `salidas/` (no se suben al repositorio).

## Instalador
`Reconstruir_Instalador.bat` genera el `.exe` (PyInstaller) y el instalador (Inno Setup, `INSTALADOR/instalador.iss`).
El instalador `Instalar_INEScan2x.exe` se distribuye en la sección **Releases**.
