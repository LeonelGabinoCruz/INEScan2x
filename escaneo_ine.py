"""
INEScan 2x - Escaneo de INE / credencial (frente y reverso) -> PDF al DOBLE de tamaño, listo para imprimir.
Pensado para HP Smart Tank 525 (usa WIA de Windows).

Flujo:
  1. Pon la credencial por el FRENTE en el cristal -> "Escanear frente".
  2. Voltéala -> "Escanear reverso".
  3. "PDF al doble (2x)" o "PDF tamaño normal" -> se recorta solo y se abre el PDF.
"""
import os
import sys
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

import numpy as np
import cv2
from PIL import Image, ImageDraw, ImageTk
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import mm

DPI = 200                         # rápido (la calidad con la que venía antes); 300 o 600 = más calidad y más lento
FACTOR = 2.0                      # tamaño final = 2x el original
INE_MM = (85.6, 54.0)           # tamaño estándar de tarjeta (ancho, alto)
MARGEN_MM = 1.0                   # borde blanco alrededor (mm reales de la credencial; en el PDF salen x2)
if getattr(sys, "frozen", False):          # instalado como .exe: guardar en Documentos
    CARPETA = os.path.join(os.path.expanduser("~"), "Documents", "INEScan 2x")
else:
    CARPETA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "salidas")
WIA_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"


# ----------------------------------------------------------------- escáner (WIA)
def listar_escaneres():
    import win32com.client
    dm = win32com.client.Dispatch("WIA.DeviceManager")
    res = []
    for i in range(1, dm.DeviceInfos.Count + 1):
        info = dm.DeviceInfos.Item(i)
        if info.Type == 1:  # 1 = escáner
            nombre = info.Properties("Name").Value
            res.append((nombre, info.DeviceID))
    return res


def _set_prop(item, pid, valor):
    """WIA: Properties(x) busca por posición, no por ID; hay que recorrer."""
    for p in item.Properties:
        if p.PropertyID == pid:
            try:
                p.Value = valor
                return True
            except Exception:
                return False
    return False


def _get_prop(item, pid):
    for p in item.Properties:
        if p.PropertyID == pid:
            return p.Value
    return None


def escanear_wia(device_id, dpi=DPI):
    """Escanea el cristal plano a color y devuelve una imagen PIL (RGB)."""
    import win32com.client
    dm = win32com.client.Dispatch("WIA.DeviceManager")
    dev = None
    for i in range(1, dm.DeviceInfos.Count + 1):
        if dm.DeviceInfos.Item(i).DeviceID == device_id:
            dev = dm.DeviceInfos.Item(i).Connect()
            break
    if dev is None:
        raise RuntimeError("No se pudo conectar con el escáner.")
    item = dev.Items.Item(1)
    _set_prop(item, 6146, 1)        # intent: color
    _set_prop(item, 4104, 24)       # 24 bits de color
    _set_prop(item, 6147, dpi)      # resolución horizontal
    _set_prop(item, 6148, dpi)      # resolución vertical
    real = _get_prop(item, 6147)
    if real != dpi:
        raise RuntimeError(f"El escáner no aceptó {dpi} dpi (quedó en {real}).")
    # se deja el área completa del cristal (por defecto)
    img = item.Transfer(WIA_BMP)
    tmp = os.path.join(tempfile.gettempdir(), "ine_scan_tmp.bmp")
    if os.path.exists(tmp):
        os.remove(tmp)
    img.SaveFile(tmp)
    im = Image.open(tmp).convert("RGB")
    im.load()
    os.remove(tmp)
    im.dpi_real = real
    return im


# ----------------------------------------------------------------- recorte automático
def recortar_credencial(pil_img):
    """Detecta la credencial, la endereza y la recorta dejando un margen blanco
    UNIFORME por los cuatro lados (fuera de la credencial todo queda blanco).
    Devuelve (imagen, ok)."""
    Image.MAX_IMAGE_PIXELS = None
    bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    h, w = bgr.shape[:2]
    esc = 800.0 / max(h, w)
    peq = cv2.resize(bgr, None, fx=esc, fy=esc, interpolation=cv2.INTER_AREA)
    gris = cv2.cvtColor(peq, cv2.COLOR_BGR2GRAY)
    gris = cv2.GaussianBlur(gris, (5, 5), 0)

    borde = np.concatenate([gris[:10].ravel(), gris[-10:].ravel(),
                            gris[:, :10].ravel(), gris[:, -10:].ravel()])
    fondo = float(np.median(borde))
    diff = cv2.absdiff(gris, np.full_like(gris, int(fondo)))
    _, mask = cv2.threshold(diff, 14, 255, cv2.THRESH_BINARY)
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k)

    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return pil_img, False
    c = max(cnts, key=cv2.contourArea)
    if cv2.contourArea(c) < 0.03 * gris.shape[0] * gris.shape[1]:
        return pil_img, False

    # silueta gruesa (incluye algo de sombra) -> sirve solo para ubicar la credencial
    sil = np.zeros(gris.shape, np.uint8)
    cv2.drawContours(sil, [c], -1, 255, thickness=cv2.FILLED)
    sil = cv2.resize(sil, (w, h), interpolation=cv2.INTER_LINEAR)

    (cx, cy), (rw, rh), ang = cv2.minAreaRect(c)
    cx, cy = cx / esc, cy / esc
    if rw < rh:
        ang += 90
    M = cv2.getRotationMatrix2D((cx, cy), ang, 1.0)
    rot = cv2.warpAffine(bgr, M, (w, h), flags=cv2.INTER_CUBIC,
                         borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))
    rsil = cv2.warpAffine(sil, M, (w, h), flags=cv2.INTER_LINEAR)
    ys, xs = np.where(rsil > 127)
    if len(xs) == 0:
        return pil_img, False
    bx0, bx1, by0, by1 = xs.min(), xs.max() + 1, ys.min(), ys.max() + 1
    bw, bh = bx1 - bx0, by1 - by0

    # --- orillas reales: una línea recta por lado, donde hay el mayor cambio
    g = cv2.cvtColor(rot, cv2.COLOR_BGR2GRAY).astype(np.float32)
    win = max(6, int(0.012 * bw))

    def borde(perfil, centro, lado):
        """perfil 1D; busca el mayor cambio cerca de 'centro' (índice grueso)."""
        perfil = cv2.GaussianBlur(perfil.reshape(1, -1), (0, 0), 1.2, sigmaY=0).ravel()
        d = np.abs(np.diff(perfil))
        lo, hi = max(0, centro - win), min(len(d), centro + win)
        if hi <= lo:
            return centro
        seg = d[lo:hi]
        # la orilla externa: primer cambio fuerte viniendo desde afuera
        fuertes = np.where(seg >= 0.45 * seg.max())[0]
        if lado in ("izq", "arriba"):
            k = fuertes.min()
        else:
            k = fuertes.max()
        return lo + int(k) + 1

    ry0, ry1 = by0 + int(0.2 * bh), by1 - int(0.2 * bh)      # franja central (alto)
    rx0, rx1 = bx0 + int(0.2 * bw), bx1 - int(0.2 * bw)      # franja central (ancho)
    tx0 = borde(g[ry0:ry1, :].mean(0), bx0, "izq")
    tx1 = borde(g[ry0:ry1, :].mean(0), bx1, "der")
    ty1 = borde(g[:, rx0:rx1].mean(1), by1, "abajo")
    # arriba: evitar la muesca central usando solo el tramo derecho
    tx_a, tx_b = bx0 + int(0.6 * bw), bx1 - int(0.08 * bw)
    ty0 = borde(g[:, tx_a:tx_b].mean(1), by0, "arriba")
    tw, th = tx1 - tx0, ty1 - ty0
    if tw < 0.5 * bw or th < 0.5 * bh:         # medición dudosa: usar caja gruesa
        tx0, tx1, ty0, ty1 = bx0, bx1, by0, by1
        tw, th = tw if False else bw, bh

    # --- máscara de esquinas redondeadas con borde suave (supermuestreo 4x)
    m = int(round(tw * (MARGEN_MM / INE_MM[0])))               # margen uniforme
    x0, y0, x1, y1 = tx0 - m, ty0 - m, tx1 + m, ty1 + m
    SS = 4
    mk = Image.new("L", ((x1 - x0) * SS, (y1 - y0) * SS), 0)
    radio = tw * (3.2 / INE_MM[0])
    ImageDraw.Draw(mk).rounded_rectangle(
        (m * SS, m * SS, (m + tw) * SS - 1, (m + th) * SS - 1), radius=radio * SS, fill=255)
    alfa = np.array(mk.resize((x1 - x0, y1 - y0), Image.LANCZOS)).astype(np.float32) / 255.0

    lienzo = np.full((y1 - y0, x1 - x0, 3), 255, np.uint8)
    sx0, sy0 = max(x0, 0), max(y0, 0)
    sx1, sy1 = min(x1, w), min(y1, h)
    lienzo[sy0 - y0:sy1 - y0, sx0 - x0:sx1 - x0] = rot[sy0:sy1, sx0:sx1]
    out = lienzo.astype(np.float32) * alfa[..., None] + 255.0 * (1 - alfa[..., None])
    res = Image.fromarray(cv2.cvtColor(out.astype(np.uint8), cv2.COLOR_BGR2RGB))
    res.dpi_real = getattr(pil_img, "dpi_real", DPI)
    res.card_px = (tw, th)           # medida de la credencial, sin el margen
    res.margin_px = m
    return res, True


# ----------------------------------------------------------------- PDF
MARGEN_SUP_NORMAL_MM = 30.0       # tamaño normal: 3 cm arriba (hasta el borde de la credencial)


def generar_pdf(frente, reverso, ruta, pagina=letter, dpi=DPI, factor=FACTOR):
    """factor=2 -> al doble, una debajo de la otra, centradas.
    factor=1 -> tamaño normal, lado a lado a la misma altura, 3 cm desde arriba."""
    pw, ph = pagina
    c = canvas.Canvas(ruta, pagesize=pagina)
    c.setTitle("Credencial %gx" % factor)
    tmpdir = tempfile.mkdtemp()
    piezas = []
    for nombre, im in (("frente", frente), ("reverso", reverso)):
        d = getattr(im, "dpi_real", dpi)
        p = os.path.join(tmpdir, nombre + ".png")
        im.save(p, dpi=(d, d))              # PNG sin pérdida
        # tamaño tomado del escaneo: píxeles / dpi, por el factor
        w = im.width / d * 25.4 * factor * mm
        h = im.height / d * 25.4 * factor * mm
        marg = getattr(im, "margin_px", 0) / d * 25.4 * factor * mm   # borde blanco
        piezas.append((p, w, h, marg))
    if factor == 1:
        gap = 10 * mm
        total = sum(p[1] for p in piezas) + gap
        x = (pw - total) / 2
        for p, w, h, marg in piezas:
            # la credencial (sin el borde blanco) queda a 3 cm del borde superior
            y = ph - MARGEN_SUP_NORMAL_MM * mm + marg - h
            c.drawImage(p, x, y, width=w, height=h)
            x += w + gap
    else:
        gap = 12 * mm
        total = sum(p[2] for p in piezas) + gap
        y = ph - (ph - total) / 2          # centrado vertical
        for p, w, h, marg in piezas:
            x = (pw - w) / 2
            y -= h
            c.drawImage(p, x, y, width=w, height=h)
            y -= gap
    c.showPage()
    c.save()



# ----------------------------------------------------------------- iconos a color
_ICO_CACHE = {}


def icono(emoji, px=28):
    """Devuelve un PhotoImage a color del emoji, del tamaño px (requiere Tk ya creado)."""
    from PIL import ImageFont
    clave = (emoji, px)
    if clave in _ICO_CACHE:
        return _ICO_CACHE[clave]
    try:
        fuente = ImageFont.truetype(r"C:\Windows\Fonts\seguiemj.ttf", 109)
        im = Image.new("RGBA", (136, 128), (0, 0, 0, 0))
        ImageDraw.Draw(im).text((4, 0), emoji, font=fuente, embedded_color=True)
        im = im.crop(im.getbbox() or (0, 0, 136, 128))
        lado = max(im.size)
        cuadro = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
        cuadro.paste(im, ((lado - im.width) // 2, (lado - im.height) // 2))
        ph = ImageTk.PhotoImage(cuadro.resize((px, px), Image.LANCZOS))
    except Exception:
        ph = None
    _ICO_CACHE[clave] = ph
    return ph


# ----------------------------------------------------------------- interfaz
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("INEScan 2x")
        ancho, alto = 800, 660
        x = (self.winfo_screenwidth() - ancho) // 2
        y = (self.winfo_screenheight() - alto) // 2
        self.geometry(f"{ancho}x{alto}+{x}+{y}")   # centrada en la pantalla
        self.resizable(False, False)               # tamaño fijo
        self.frente = None
        self.reverso = None
        self.escaneres = []

        st = ttk.Style(self)
        st.configure("Icono.TButton", padding=(10, 6))

        def boton(padre, emoji, texto, cmd, px=26):
            return ttk.Button(padre, text=" " + texto, image=icono(emoji, px),
                              compound="left", command=cmd, style="Icono.TButton")

        def etiqueta(padre, emoji, texto, px=24):
            return ttk.Label(padre, text=" " + texto, image=icono(emoji, px), compound="left")

        cab = ttk.Frame(self, padding=(12, 10, 12, 0))
        cab.pack(fill="x")
        ttk.Label(cab, text=" INEScan 2x", image=icono("🪪", 40), compound="left",
                  font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Label(cab, text="Escanea tu credencial y obtén el PDF listo para imprimir",
                  foreground="#555").pack(side="left", padx=14, pady=(8, 0))

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        etiqueta(top, "🖨️", "Escáner:").pack(side="left")
        self.cmb = ttk.Combobox(top, state="readonly", width=40)
        self.cmb.pack(side="left", padx=6)
        boton(top, "🔄", "Buscar", self.buscar).pack(side="left")
        etiqueta(top, "📄", "Papel:").pack(side="left", padx=(14, 4))
        self.papel = ttk.Combobox(top, state="readonly", width=7, values=["Carta", "A4"])
        self.papel.current(0)
        self.papel.pack(side="left")

        mid = ttk.Frame(self, padding=10)
        mid.pack(fill="both", expand=True)
        self.lbl = {}
        for i, (clave, emoji, texto) in enumerate((("frente", "🪪", "FRENTE"), ("reverso", "🔁", "REVERSO"))):
            f = ttk.LabelFrame(mid, labelwidget=etiqueta(mid, emoji, texto, 28), padding=6)
            f.grid(row=0, column=i, padx=8, sticky="nsew")
            mid.columnconfigure(i, weight=1)
            l = ttk.Label(f, text="(sin escanear)", image=icono("🖼️", 72), compound="top", anchor="center")
            l.pack(fill="both", expand=True)
            self.lbl[clave] = l
            boton(f, "📷", "Escanear " + texto.lower(),
                  lambda c=clave: self.escanear(c)).pack(pady=4, fill="x")
            boton(f, "📂", "Abrir imagen…",
                  lambda c=clave: self.abrir(c)).pack(fill="x")
        mid.rowconfigure(0, weight=1)

        bot = ttk.Frame(self, padding=10)
        bot.pack(fill="x")
        self.ico_estado = ttk.Label(bot)
        self.ico_estado.pack(side="left")
        self.estado = ttk.Label(bot, text="")
        self.estado.pack(side="left", padx=6)
        boton(bot, "📄", "PDF tamaño normal", lambda: self.pdf(1)).pack(side="right")
        boton(bot, "🔍", "PDF al doble (2x)", lambda: self.pdf(2)).pack(side="right", padx=6)
        boton(bot, "🧹", "Limpiar", self.limpiar).pack(side="right", padx=6)

        self.msg("ℹ️ Pon la credencial en el cristal, esquina superior, y escanea el FRENTE.")
        self.after(200, self.buscar)

    def msg(self, t):
        """El primer símbolo del mensaje (✅ ⚠️ ❌ ⏳ ℹ️ ...) se muestra como icono grande."""
        t = t.strip()
        icono_txt = ""
        for e in ("✅", "⚠️", "❌", "⏳", "ℹ️", "🧹"):
            if t.startswith(e):
                icono_txt, t = e, t[len(e):].strip()
                break
        t = t.replace("🔁 ", "").replace("🪪 ", "")
        self.ico_estado.config(image=icono(icono_txt, 30) if icono_txt else "")
        self.estado.config(text=t, wraplength=250, justify="left")
        self.update_idletasks()

    def buscar(self):
        try:
            self.escaneres = listar_escaneres()
        except Exception as e:
            self.escaneres = []
            self.msg(f"❌ Error buscando escáner: {e}")
        self.cmb["values"] = [n for n, _ in self.escaneres]
        if self.escaneres:
            # preferir HP
            idx = next((i for i, (n, _) in enumerate(self.escaneres) if "HP" in n.upper() or "SMART" in n.upper()), 0)
            self.cmb.current(idx)
            self.msg("✅ Escáner listo.")
        else:
            self.msg("⚠️ No se encontró escáner. Enciende la impresora/conéctala o usa 'Abrir imagen…'.")

    def mostrar(self, clave, im):
        mini = im.copy()
        mini.thumbnail((320, 230))
        ph = ImageTk.PhotoImage(mini)
        self.lbl[clave].config(image=ph, text="", compound="center")
        self.lbl[clave].image = ph

    def procesar(self, clave, im):
        rec, ok = recortar_credencial(im)
        setattr(self, clave, rec)
        self.mostrar(clave, rec)
        if not ok:
            self.msg(f"⚠️ {clave}: no detecté bien la credencial, se usó la imagen completa. Reescanea sobre fondo claro.")
        elif clave == "frente":
            self.msg("✅ Frente listo. 🔁 Voltea la credencial y escanea el REVERSO.")
        else:
            self.msg("✅ Reverso listo. Elige 'PDF al doble (2x)' o 'PDF tamaño normal'.")

    def escanear(self, clave):
        i = self.cmb.current()
        if i < 0:
            messagebox.showwarning("⚠️ Escáner", "⚠️ No hay escáner seleccionado.\nPresiona 🔄 Buscar o usa 📂 Abrir imagen.")
            return
        dev_id = self.escaneres[i][1]
        self.msg(f"⏳ Escaneando {clave}… espera.")

        def trabajo():
            try:
                import pythoncom
                pythoncom.CoInitialize()
                im = escanear_wia(dev_id)
                self.after(0, lambda: self.procesar(clave, im))
            except Exception as e:
                self.after(0, lambda: (self.msg("❌ Falló el escaneo."),
                                       messagebox.showerror("❌ Escaneo", f"❌ No se pudo escanear:\n{e}")))
        threading.Thread(target=trabajo, daemon=True).start()

    def abrir(self, clave):
        ruta = filedialog.askopenfilename(filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff")])
        if ruta:
            src = Image.open(ruta)
            d = src.info.get("dpi", (DPI, DPI))[0] or DPI
            im = src.convert("RGB")
            im.dpi_real = d          # si el archivo no trae dpi, se asume DPI
            self.procesar(clave, im)

    def limpiar(self):
        self.frente = self.reverso = None
        for l in self.lbl.values():
            l.config(image=icono("🖼️", 72), text="(sin escanear)", compound="top")
        self.msg("🧹 Listo. Escanea el FRENTE.")

    def pdf(self, factor=2):
        if self.frente is None or self.reverso is None:
            messagebox.showinfo("ℹ️ Falta algo", "ℹ️ Escanea primero 🪪 frente y 🔁 reverso.")
            return
        os.makedirs(CARPETA, exist_ok=True)
        ruta = os.path.join(CARPETA, datetime.now().strftime("credencial_%Y%m%d_%H%M%S") + ("_normal.pdf" if factor == 1 else "_2x.pdf"))
        pagina = letter if self.papel.get() == "Carta" else A4
        try:
            generar_pdf(self.frente, self.reverso, ruta, pagina, factor=factor)
        except Exception as e:
            messagebox.showerror("❌ PDF", "❌ No se pudo crear el PDF:\n" + str(e))
            return
        self.msg(f"✅ PDF guardado: {ruta}")
        os.startfile(ruta)


if __name__ == "__main__":
    App().mainloop()
