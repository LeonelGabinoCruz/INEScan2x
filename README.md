# INEScan 2x

Escanea credenciales y documentos con un escáner de Windows (WIA, probado con HP Smart Tank 525) y genera PDF listos para imprimir.

Todo funciona en una sola ventana. Al abrir la app aparece el **inicio**: el título, filtros por categoría (**Todos**, **Escanear**, **Optimizar PDF**) y una tarjeta por cada función: **Credencial (INE)**, **Documento oficio**, **Varias páginas** y **Comprimir PDF**. Al elegir una se muestra en la misma ventana, con el selector de escáner y papel arriba (en las que escanean) y un botón **Regresar** para volver al inicio. La barra superior está siempre visible: el logo regresa al inicio y sus enlaces llevan directo a cada función. Cada función conserva lo que ya escaneaste mientras la app esté abierta.

## Credencial (INE)
Escanea frente y reverso de una INE (o cualquier credencial):

- **PDF al doble (2x):** la credencial al doble de su tamaño real, una debajo de la otra.
- **PDF tamaño normal:** frente y reverso lado a lado, a 3 cm del borde superior.

Recorta y endereza solo, con margen blanco uniforme. El tamaño se toma de las medidas reales del escaneo.

## Documento oficio en 2 partes
Para documentos más largos que el cristal del escáner o que la hoja de tu impresora (p. ej. un acta en tamaño oficio). En el inicio, elige **Documento oficio**:

1. Escanea la parte de **arriba**.
2. Desliza el documento y escanea lo que falta, dejando ~2 cm repetidos de lo ya escaneado.
3. La app detecta dónde continúa el documento (puntos de referencia), alinea las dos partes, empareja el brillo y las une.
4. Detecta lo que realmente es papel y recorta el cristal vacío; endereza la inclinación y deja los bordes rectos. Con papel de color usa el color del papel; con papel blanco recorta al contenido (margen de 10 mm).
5. Elige la forma de impresión y genera el PDF:
   - **Oficio, tamaño real** (21.6 x 34 cm)
   - **Ajustar a 1 hoja carta/A4** (reducido)
   - **Dividir en hojas carta/A4** (tamaño real, cortando por una fila lo más vacía posible)

## Varias páginas en un solo PDF
Para juntar varias hojas o imágenes en un solo PDF (p. ej. un trámite con varios documentos). En el inicio, elige **Varias páginas**:

1. Pon una hoja en el cristal y presiona **Escanear página**; repite con cada hoja. También puedes usar **Agregar imágenes…** (varias a la vez; los TIFF de varias páginas se agregan completos y las fotos del celular se enderezan solas).
2. Ordena las páginas con **Subir / Bajar**, gíralas 90° con **Girar** o elimínalas con **Quitar**.
3. La hoja es la elegida en **Papel** (Carta o A4); marca **Hoja del tamaño de cada imagen** si prefieres que cada hoja mida lo mismo que su imagen. Elige la **compresión**:
   - **Sin comprimir:** máxima calidad, sin pérdida (archivo grande).
   - **Normal (recomendado):** 150 dpi, buena calidad.
   - **Alta:** 120 dpi, archivo pequeño.
   - **Máxima:** 96 dpi, lo más pequeño posible (para portales con límite de tamaño).
   - **Escala de grises:** lo reduce aún más.
4. **Generar PDF**: se guarda en `salidas/`, se abre solo y se muestra su tamaño final.

## Comprimir PDF
Para hacer más ligero un PDF que ya tienes (p. ej. para subirlo a un portal con límite de tamaño). En el inicio, elige **Comprimir PDF**:

1. **Elegir PDF…** muestra su tamaño actual y número de páginas.
2. Mueve la barra al porcentaje de su tamaño actual que quieres que pese (de 10% a 90%, de 5 en 5); se muestra cuánto pesaría.
3. **Comprimir**: la app recomprime las imágenes del PDF (el texto y los dibujos no se tocan) y busca la mejor calidad que cumpla el porcentaje pedido. Si basta con limpiar el archivo, lo hace sin pérdida.
4. Se muestra **antes → ahora**, cuánto se redujo y la calidad resultante. La copia se guarda junto al original como `<nombre>_comprimido_<porcentaje>.pdf`; **el original no se modifica**.
5. Abajo aparece el **historial de ese archivo**: cada compresión con fecha, tamaño antes y después, porcentaje pedido y calidad. Incluye las copias comprimidas que salen de él (si comprimes una copia, su historial se muestra junto con el del original).
6. Elige cualquier versión de la lista y usa **Guardar como…** para guardarla donde quieras, **Abrir** para verla (o doble clic) o **Quitar del historial**. La app guarda su propia copia de cada versión (en `salidas/versiones_pdf/`, o en `Documentos\INEScan 2x\versiones_pdf\` con el instalador), así que cualquier tamaño se puede volver a guardar aunque hayas borrado o movido el archivo de junto al original. Quitar una versión del historial borra esa copia interna, pero no el archivo de junto al original.

Un PDF de solo texto casi no se puede reducir (ya es ligero); uno protegido con contraseña no se puede comprimir.

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
