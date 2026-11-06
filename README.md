# INEScan 2x

Escanea frente y reverso de una INE (o cualquier credencial) con un escáner de Windows (WIA, probado con HP Smart Tank 525) y genera un PDF listo para imprimir:

- **PDF al doble (2x):** la credencial al doble de su tamaño real, una debajo de la otra.
- **PDF tamaño normal:** frente y reverso lado a lado, a 3 cm del borde superior.

Recorta y endereza solo, con margen blanco uniforme. El tamaño se toma de las medidas reales del escaneo.

## Uso
```
pip install -r requirements.txt
python escaneo_ine.py
```
o doble clic en `Abrir_Escaneo.bat`. Los PDF se guardan en `salidas/`.

## Instalador
`Reconstruir_Instalador.bat` genera el `.exe` (PyInstaller) y el instalador (Inno Setup, `INSTALADOR/instalador.iss`).
El instalador `Instalar_INEScan2x.exe` se distribuye en la sección **Releases**.
