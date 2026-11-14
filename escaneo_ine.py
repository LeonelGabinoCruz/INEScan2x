"""
INEScan 2x - Escaneo de INE / credencial (frente y reverso) -> PDF al DOBLE de tamaño, listo para imprimir.
Pensado para HP Smart Tank 525 (usa WIA de Windows).

Una sola ventana: el inicio muestra el título y las funciones; cada función se abre en la misma ventana
con un botón "Regresar".
  - Credencial (INE): escanea frente y reverso -> "PDF al doble (2x)" o "PDF tamaño normal".
  - Documento oficio: documento largo escaneado en 2 partes que se unen solas.
  - Varias páginas: varias hojas/imágenes en un solo PDF, con compresión opcional.
  - Comprimir PDF: reduce un PDF existente al porcentaje elegido, con historial por archivo.
"""
import json
import os
import shutil
import sys
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
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
CONFIG = os.path.join(CARPETA, "config.json")
ICONO_APP = os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))), "icono.ico")


def cargar_config():
    try:
        with open(CONFIG, encoding="utf8") as f:
            cfg = json.load(f)
        return cfg if isinstance(cfg, dict) else {}
    except Exception:
        return {}


def guardar_config(cfg):
    try:
        os.makedirs(CARPETA, exist_ok=True)
        with open(CONFIG, "w", encoding="utf8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass                                       # no es crítico


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


WIA_FORMATOS = [                                    # en orden de preferencia (sin pérdida primero)
    "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}",       # BMP
    "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}",       # PNG
    "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}",       # TIFF
    "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}",       # JPEG
]


def _opciones(item, pid):
    """Lo que el driver acepta para una propiedad: ('lista', [..]) o ('rango', (min, max, paso)) o None."""
    for p in item.Properties:
        if p.PropertyID == pid:
            try:
                if p.SubType == 2:
                    return "lista", [int(v) for v in p.SubTypeValues]
                if p.SubType == 1:
                    return "rango", (int(p.SubTypeMin), int(p.SubTypeMax), max(1, int(p.SubTypeStep)))
            except Exception:
                return None
    return None


def _dpi_soportado(item, deseado):
    """Resolución que el escáner realmente ofrece más cercana a la deseada (no todos aceptan 200)."""
    op = _opciones(item, 6147)
    if not op:
        return deseado
    tipo, v = op
    if tipo == "lista":
        return min(v, key=lambda x: (abs(x - deseado), -x))
    lo, hi, paso = v
    d = min(max(deseado, lo), hi)
    return lo + round((d - lo) / paso) * paso


def _elegir_item(dev):
    """El elemento del cristal plano (si el equipo tiene varios, p. ej. alimentador automático)."""
    n = dev.Items.Count
    if n == 1:
        return dev.Items.Item(1)
    for i in range(1, n + 1):
        it = dev.Items.Item(i)
        nombre = str(_get_prop(it, 4098) or "").lower()
        if any(k in nombre for k in ("flat", "plano", "cristal", "platen", "glass")):
            return it
    return dev.Items.Item(1)


def escanear_wia(device_id, dpi=DPI):
    """Escanea el cristal plano a color con WIA (driver estándar de Windows: HP, Epson, Canon, Brother...)
    y devuelve una imagen PIL (RGB). Se adapta a lo que cada driver ofrece."""
    import win32com.client
    dm = win32com.client.Dispatch("WIA.DeviceManager")
    dev = None
    for i in range(1, dm.DeviceInfos.Count + 1):
        if dm.DeviceInfos.Item(i).DeviceID == device_id:
            dev = dm.DeviceInfos.Item(i).Connect()
            break
    if dev is None:
        raise RuntimeError("No se pudo conectar con el escáner. ¿Está encendido y conectado?")
    caps = _get_prop(dev, 3086)                       # capacidades: 1 = alimentador, 2 = cristal plano
    if isinstance(caps, int) and caps & 2:
        _set_prop(dev, 3088, 2)                       # usar el cristal plano
    item = _elegir_item(dev)
    _set_prop(item, 6146, 1)                          # intento: color
    _set_prop(item, 4104, 24)                         # 24 bits de color
    objetivo = _dpi_soportado(item, dpi)
    _set_prop(item, 6147, objetivo)
    _set_prop(item, 6148, objetivo)
    real = _get_prop(item, 6147) or objetivo
    _set_prop(item, 6149, 0)                          # toda el área del cristal
    _set_prop(item, 6150, 0)
    for pid in (6151, 6152):
        op = _opciones(item, pid)
        if op and op[0] == "rango":
            _set_prop(item, pid, op[1][1])

    try:
        disponibles = {str(f) for f in item.Formats}
    except Exception:
        disponibles = set()
    orden = [f for f in WIA_FORMATOS if not disponibles or f in disponibles] or WIA_FORMATOS
    img, ultimo = None, None
    for fmt in orden + [None]:                        # al final, el formato nativo del driver
        try:
            img = item.Transfer(fmt) if fmt else item.Transfer()
            break
        except Exception as e:
            ultimo = e
    if img is None:
        raise RuntimeError(f"El escáner no entregó la imagen: {ultimo}")
    ext = "." + str(getattr(img, "FileExtension", "bmp") or "bmp").lstrip(".")
    tmp = os.path.join(tempfile.gettempdir(), "ine_scan_tmp" + ext)
    if os.path.exists(tmp):
        os.remove(tmp)
    img.SaveFile(tmp)
    im = Image.open(tmp).convert("RGB")
    im.load()
    os.remove(tmp)
    try:
        dpi_img = int(img.HorizontalResolution)
    except Exception:
        dpi_img = 0
    # el encabezado de la imagen redondea (199.99 -> 199); se usa el valor fijado en el driver
    # salvo que el driver haya remuestreado y difiera claramente
    im.dpi_real = real if (dpi_img <= 0 or abs(real - dpi_img) <= 0.03 * real) else dpi_img
    return im


# ----------------------------------------------------------------- escáner de red (eSCL / AirScan)
def _escl_http(url, datos=None, metodo=None, timeout=30):
    import ssl
    import urllib.request
    req = urllib.request.Request(url, data=datos, method=metodo,
                                headers={"Content-Type": "text/xml"} if datos else {})
    ctx = ssl._create_unverified_context()           # las impresoras usan certificado propio
    return urllib.request.urlopen(req, timeout=timeout, context=ctx)


def _escl_base(host):
    """Dirección base que responde a eSCL (prueba http/https y los puertos habituales)."""
    host = host.strip().rstrip("/")
    if host.startswith("http"):
        cands = [host]
    elif ":" in host:
        cands = [f"http://{host}", f"https://{host}"]
    else:
        cands = [f"http://{host}", f"https://{host}", f"http://{host}:8080"]
    ultimo = None
    for base in cands:
        try:
            with _escl_http(base + "/eSCL/ScannerCapabilities", timeout=6) as r:
                return base, r.read()
        except Exception as e:
            ultimo = e
    raise RuntimeError(f"No encontré un escáner de red en '{host}' (eSCL/AirScan). {ultimo}")


def escanear_escl(host, dpi=DPI):
    """Escanea el cristal plano de una impresora de red por eSCL (AirScan/Mopria), sin drivers."""
    import io
    import time
    import xml.etree.ElementTree as ET
    base, xml_caps = _escl_base(host)
    raiz = ET.fromstring(xml_caps)
    for el in raiz.iter():
        el.tag = el.tag.split("}")[-1]               # sin espacios de nombres
    plat = raiz.find(".//Platen/PlatenInputCaps")
    if plat is None:
        raise RuntimeError("Ese escáner de red no tiene cristal plano.")
    ancho = int(plat.findtext("MaxWidth") or 2550)    # en 1/300 de pulgada
    alto = int(plat.findtext("MaxHeight") or 3507)
    res = sorted({int(x.text) for x in plat.iter("XResolution") if x.text and x.text.strip().isdigit()})
    dpi_u = min(res, key=lambda x: (abs(x - dpi), -x)) if res else dpi
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<scan:ScanSettings xmlns:pwg="http://www.pwg.org/schemas/2010/12/sm" '
        'xmlns:scan="http://schemas.hp.com/imaging/escl/2011/05/03">'
        '<pwg:Version>2.0</pwg:Version><scan:Intent>Document</scan:Intent>'
        '<pwg:ScanRegions><pwg:ScanRegion>'
        '<pwg:ContentRegionUnits>escl:ThreeHundredthsOfInches</pwg:ContentRegionUnits>'
        '<pwg:XOffset>0</pwg:XOffset><pwg:YOffset>0</pwg:YOffset>'
        f'<pwg:Width>{ancho}</pwg:Width><pwg:Height>{alto}</pwg:Height>'
        '</pwg:ScanRegion></pwg:ScanRegions>'
        '<pwg:InputSource>Platen</pwg:InputSource><scan:ColorMode>RGB24</scan:ColorMode>'
        f'<scan:XResolution>{dpi_u}</scan:XResolution><scan:YResolution>{dpi_u}</scan:YResolution>'
        '<pwg:DocumentFormat>image/jpeg</pwg:DocumentFormat></scan:ScanSettings>'
    )
    try:
        with _escl_http(base + "/eSCL/ScanJobs", xml.encode("utf-8"), "POST") as r:
            ubic = r.headers.get("Location")
    except Exception as e:
        raise RuntimeError(f"El escáner de red rechazó el trabajo: {e}")
    if not ubic:
        raise RuntimeError("El escáner de red no devolvió la ubicación del trabajo.")
    if ubic.startswith("/"):
        ubic = base + ubic
    elif not ubic.startswith("http"):
        ubic = base + "/eSCL/ScanJobs/" + ubic
    fin = time.time() + 120
    while True:
        try:
            with _escl_http(ubic.rstrip("/") + "/NextDocument", timeout=90) as r:
                datos = r.read()
            break
        except Exception as e:
            if time.time() > fin or "503" not in str(e):      # 503 = aún ocupado: reintentar
                raise RuntimeError(f"No se recibió la imagen del escáner de red: {e}")
            time.sleep(1.5)
    im = Image.open(io.BytesIO(datos)).convert("RGB")
    im.load()
    im.dpi_real = dpi_u
    return im


def escanear(device_id, dpi=DPI):
    """Escanea con el dispositivo elegido: WIA (USB/Windows) o eSCL ('escl:<ip>', red)."""
    if device_id.startswith("escl:"):
        return escanear_escl(device_id[5:], dpi)
    return escanear_wia(device_id, dpi)


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


# ----------------------------------------------------------------- documento oficio en 2 partes
OFICIO = (215.9 / 25.4 * 72, 340.0 / 25.4 * 72)     # 21.6 x 34 cm en puntos
UMBRAL_PLANTILLA = 0.60                              # coincidencia mínima del método de respaldo
MIN_PUNTOS = 15                                      # puntos de referencia mínimos para aceptar la unión


def _gris(im):
    return cv2.cvtColor(np.array(im.convert("RGB")), cv2.COLOR_RGB2GRAY)


def _filas_con_contenido(g, umbral=235):
    return (g < umbral).sum(1) > max(3, int(0.002 * g.shape[1]))


def _alinear_sift(g1, g2):
    """Puntos de referencia (SIFT) + RANSAC: matriz 2x3 que lleva la parte 2 al marco de la parte 1
    (desplazamiento + pequeño giro). Devuelve (M, puntos) o (None, 0)."""
    s = min(1.0, 1000.0 / max(g1.shape[1], g2.shape[1]))
    cl = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    a = cl.apply(cv2.resize(g1, None, fx=s, fy=s, interpolation=cv2.INTER_AREA))
    b = cl.apply(cv2.resize(g2, None, fx=s, fy=s, interpolation=cv2.INTER_AREA))
    sift = cv2.SIFT_create(nfeatures=5000)
    k1, d1 = sift.detectAndCompute(a, None)
    k2, d2 = sift.detectAndCompute(b, None)
    if d1 is None or d2 is None or len(k1) < 10 or len(k2) < 10:
        return None, 0
    pares = cv2.BFMatcher().knnMatch(d2, d1, k=2)
    buenos = [p[0] for p in pares if len(p) == 2 and p[0].distance < 0.75 * p[1].distance]
    if len(buenos) < 10:
        return None, 0
    src = np.float32([k2[m.queryIdx].pt for m in buenos])
    dst = np.float32([k1[m.trainIdx].pt for m in buenos])
    M, inl = cv2.estimateAffinePartial2D(src, dst, method=cv2.RANSAC, ransacReprojThreshold=3.0)
    if M is None:
        return None, 0
    escala = float(np.hypot(M[0, 0], M[1, 0]))
    giro = abs(float(np.degrees(np.arctan2(M[1, 0], M[0, 0]))))
    if not (0.93 < escala < 1.07) or giro > 6:            # no es el mismo documento / alineación absurda
        return None, 0
    M[:, 2] /= s                                         # a la resolución original
    return M, int(inl.sum())


def _alinear_plantilla(g1, g2, d):
    """Respaldo: busca una franja del inicio de la parte 2 dentro de la parte 1 (solo desplazamiento)."""
    f2 = _filas_con_contenido(g2)
    top2 = int(np.argmax(f2))
    bot2 = len(f2) - int(np.argmax(f2[::-1]))
    W1, W2 = g1.shape[1], g2.shape[1]
    th = int(d / 25.4 * 18)
    m = int(0.06 * W2)
    mejor = (-1.0, 0, 0, 0)
    for k in range(0, 8):
        t0 = top2 + k * th // 2
        if t0 + th > bot2:
            break
        tpl = g2[t0:t0 + th, m:W2 - m]
        if tpl.std() < 8 or tpl.shape[1] > W1 or tpl.shape[0] > g1.shape[0]:
            continue
        _, v, _, loc = cv2.minMaxLoc(cv2.matchTemplate(g1, tpl, cv2.TM_CCOEFF_NORMED))
        if v > mejor[0]:
            mejor = (v, t0, loc[0], loc[1])
    v, t0, mx, my = mejor
    if v < UMBRAL_PLANTILLA:
        return None, 0.0
    return np.float32([[1, 0, mx - m], [0, 1, my - t0]]), float(v)


def _pendiente(ys, v):
    """Pendiente robusta (px de borde por fila) de un borde casi vertical."""
    ys = ys.astype(np.float64)
    v = v.astype(np.float64)
    keep = np.ones(len(ys), bool)
    p = np.polyfit(ys, v, 1)
    for _ in range(3):
        res = v - np.polyval(p, ys)
        lim = max(3.0, 3 * np.median(np.abs(res[keep])))
        keep = np.abs(res) <= lim
        if keep.sum() < 20:
            break
        p = np.polyfit(ys[keep], v[keep], 1)
    return float(p[0])


def _recorte_por_contenido(arr, d, motivo):
    """Papel blanco (no se distingue del cristal): se recorta al contenido con un margen de 10 mm."""
    ink = arr.min(2) < 200
    h, w = ink.shape
    fil = np.where(ink.sum(1) > max(3, 0.002 * w))[0]
    col = np.where(ink.sum(0) > max(3, 0.002 * h))[0]
    if len(fil) == 0 or len(col) == 0:
        return arr, "no se detectó el documento, se dejó completo"
    m = int(d / 25.4 * 10)
    y0, y1 = max(0, fil[0] - m), min(h, fil[-1] + 1 + m)
    x0, x1 = max(0, col[0] - m), min(w, col[-1] + 1 + m)
    return arr[y0:y1, x0:x1], f"{motivo}: recortado al contenido (margen de 10 mm)"


def _mascara_papel(arr):
    """Píxeles de papel de color: el cristal/tapa es gris neutro (sin saturación) y el papel verde, amarillo,
    etc. sí tiene color; la tinta oscura también cuenta. Es más fiable que un umbral de brillo."""
    a = arr.astype(np.int16)
    return ((a.max(2) - a.min(2)) > 6) | (a.min(2) < 200)


def recortar_documento(arr, d):
    """Detecta lo que realmente es papel y quita el cristal vacío escaneado alrededor.
    Papel de color: usa su color, corrige la inclinación y recorta al rectángulo de papel
    (bordes rectos). Papel blanco: recorta al contenido. Devuelve (arr, texto)."""
    mm = 25.4 / d
    cov = _mascara_papel(arr)
    if cov.mean() < 0.15:
        return _recorte_por_contenido(arr, d, "papel blanco")
    h, w = cov.shape

    def zona_papel(cov):
        """Filas y columnas donde hay papel de color (la cobertura se mide dentro del papel, no del cristal)."""
        ys = np.where(cov.mean(1) > 0.3)[0]
        if len(ys) < 0.2 * h:
            return None
        cols = np.where(cov[ys].mean(0) > 0.5)[0]
        if len(cols) < 0.2 * w:
            return None
        return ys, int(cols[0]), int(cols[-1]) + 1

    z = zona_papel(cov)
    if z is None:
        return _recorte_por_contenido(arr, d, "papel no detectado")
    ys, cx0, cx1 = z
    c = cov[ys]
    izq = np.argmax(c, axis=1)
    der = w - np.argmax(c[:, ::-1], axis=1)
    ang = float(np.degrees(np.arctan((_pendiente(ys, izq) + _pendiente(ys, der)) / 2)))
    txt_ang = ""
    if 0.05 < abs(ang) <= 6:                              # enderezar
        M = cv2.getRotationMatrix2D((w / 2, h / 2), -ang, 1.0)
        arr = cv2.warpAffine(arr, M, (w, h), flags=cv2.INTER_CUBIC, borderValue=(243, 243, 243))
        cov = _mascara_papel(arr)
        txt_ang = f", inclinación de {abs(ang):.2f}° corregida"
        z = zona_papel(cov)
        if z is None:
            return _recorte_por_contenido(arr, d, "papel no detectado")
        ys, cx0, cx1 = z
    ancho = cx1 - cx0
    n0, n1 = cx0 + int(0.05 * ancho), cx1 - int(0.05 * ancho)           # núcleo del papel
    fr = np.where(cov[:, n0:n1].mean(1) > 0.9)[0]
    if len(fr) < 0.2 * h:
        return _recorte_por_contenido(arr, d, "papel no detectado")
    y0, y1 = int(fr[0]), int(fr[-1]) + 1
    cc = np.where(cov[y0:y1].mean(0) >= 0.97)[0]
    if len(cc) < 0.2 * w:
        return _recorte_por_contenido(arr, d, "papel no detectado")
    x0, x1 = int(cc[0]), int(cc[-1]) + 1
    ins = max(1, int(1.0 / mm))                           # 1 mm hacia adentro: fuera sombra/orilla
    if x1 - x0 > 4 * ins and y1 - y0 > 4 * ins:
        x0, x1, y0, y1 = x0 + ins, x1 - ins, y0 + ins, y1 - ins
    res = arr[y0:y1, x0:x1]
    cr = _mascara_papel(res)                                 # verificación de bordes tras el recorte
    fi = cr.mean(1) > 0.5
    irreg = 0.0
    if fi.any():
        L = np.argmax(cr[fi], axis=1)
        R = cr.shape[1] - np.argmax(cr[fi][:, ::-1], axis=1)
        irreg = max(np.percentile(L, 98) - np.percentile(L, 2),
                    np.percentile(R, 98) - np.percentile(R, 2)) * mm
    estado = "bordes rectos" if irreg < 1.0 else f"⚠️ bordes con {irreg:.1f} mm de irregularidad"
    return res, (f"papel de color detectado: {res.shape[1] * mm / 10:.1f} x {res.shape[0] * mm / 10:.1f} cm"
                 f"{txt_ang}, {estado}")


def unir_partes(p1, p2):
    """Une dos escaneos del mismo documento (parte de arriba y parte de abajo).
    Detecta dónde continúa el documento, corrige desplazamiento/giro, empareja el brillo y mezcla la unión.
    Devuelve (imagen_unida, info); info=None si no encontró traslape (se apilan sin empalme)."""
    d = getattr(p1, "dpi_real", DPI)
    a1 = np.array(p1.convert("RGB"))
    a2 = np.array(p2.convert("RGB"))
    g1, g2 = _gris(p1), _gris(p2)
    f1, f2 = _filas_con_contenido(g1), _filas_con_contenido(g2)
    if not f1.any() or not f2.any():
        raise RuntimeError("Uno de los escaneos está en blanco.")
    bot1 = len(f1) - int(np.argmax(f1[::-1]))
    top2 = int(np.argmax(f2))
    bot2 = len(f2) - int(np.argmax(f2[::-1]))
    h1, w1 = g1.shape
    h2, w2 = g2.shape

    M, info = _alinear_sift(g1, g2)
    if M is not None and info >= MIN_PUNTOS:
        info = f"{info} puntos de referencia"
    else:
        M, v = _alinear_plantilla(g1, g2, d)
        info = f"coincidencia {v:.0%}" if M is not None else None

    if M is None:                                         # sin traslape: se apilan
        lienzo = np.full((bot1 + bot2 - top2, max(w1, w2), 3), 255, np.uint8)
        lienzo[:bot1, :w1] = a1[:bot1]
        lienzo[bot1:, :w2] = a2[top2:bot2]
        res = Image.fromarray(lienzo)
        res.dpi_real = d
        return res, None

    esq = np.float32([[0, 0], [w2, 0], [w2, bot2], [0, bot2]])
    esq = esq @ M[:, :2].T + M[:, 2]
    y2top, y2bot = float(esq[:, 1].min()), float(esq[:, 1].max())
    H = int(np.ceil(max(bot1, y2bot)))
    blanco = (255, 255, 255)
    w2img = cv2.warpAffine(a2, M, (w1, H), flags=cv2.INTER_LINEAR, borderValue=blanco)
    mask = np.zeros((h2, w2), np.uint8)
    mask[:bot2] = 255
    mask = cv2.warpAffine(mask, M, (w1, H), flags=cv2.INTER_NEAREST, borderValue=0) > 0
    base = np.full((H, w1, 3), 255, np.uint8)
    base[:min(bot1, H)] = a1[:min(bot1, H)]

    ya, yb = max(0, int(y2top)), min(bot1, H)             # filas donde existen las dos partes
    ov = yb - ya
    ys = np.arange(H, dtype=np.float32)[:, None]
    if ov > 8:
        zona = mask[ya:yb] & (base[ya:yb].min(2) < 255)
        if zona.sum() > 1000:                             # igualar brillo de la parte 2 a la parte 1
            ganancia = base[ya:yb][zona].mean(0) / np.maximum(w2img[ya:yb][zona].mean(0), 1.0)
            w2img = np.clip(w2img.astype(np.float32) * np.clip(ganancia, 0.85, 1.15), 0, 255)
        centro, medio = (ya + yb) / 2.0, max(2.0, 0.2 * ov)
        rampa = np.clip((ys - (centro - medio)) / (2 * medio), 0, 1)   # 0 = parte 1, 1 = parte 2
    else:
        rampa = (ys >= yb).astype(np.float32)
    rampa = np.where(ys >= bot1, 1.0, rampa)
    w = rampa * mask
    out = base.astype(np.float32) * (1 - w[..., None]) + w2img.astype(np.float32) * w[..., None]
    out = np.clip(out, 0, 255).astype(np.uint8)
    out, txt_doc = recortar_documento(np.ascontiguousarray(out), d)
    res = Image.fromarray(np.ascontiguousarray(out))
    res.dpi_real = d
    return res, f"{info}; {txt_doc}"


def generar_pdf_documento(img, ruta, modo, pagina=letter):
    """modo: 'oficio' (hoja 21.6x34 cm, tamaño real), 'carta' (reducido para caber en 1 hoja),
    'dividir' (tamaño real repartido en varias hojas carta/A4)."""
    d = getattr(img, "dpi_real", DPI)
    wpt = img.width / d * 72.0                   # tamaño real en puntos
    hpt = img.height / d * 72.0
    tmpdir = tempfile.mkdtemp()

    def guardar(im, nombre):
        p = os.path.join(tmpdir, nombre)
        im.save(p, dpi=(d, d))
        return p

    if modo in ("oficio", "carta"):
        pw, ph = OFICIO if modo == "oficio" else pagina
        s = min(1.0, pw / wpt, ph / hpt)
        w, h = wpt * s, hpt * s
        c = canvas.Canvas(ruta, pagesize=(pw, ph))
        c.setTitle("Documento")
        x = (pw - w) / 2
        c.drawImage(guardar(img, "doc.png"), x, ph - h, width=w, height=h)
        c.showPage()
        c.save()
        return

    pw, ph = pagina
    s = min(1.0, pw / wpt)
    pt_px = 72.0 / d * s                         # puntos del PDF por píxel
    filas = int(ph / pt_px)
    g = _gris(img)
    tinta = (g < 200).sum(1)
    c = canvas.Canvas(ruta, pagesize=pagina)
    c.setTitle("Documento")
    y = 0
    n = 0
    while y < img.height:
        fin = min(y + filas, img.height)
        if fin < img.height:                     # cortar por una fila lo más vacía posible
            ini = max(y + 1, fin - int(0.08 * filas))
            fin = ini + int(np.argmin(tinta[ini:fin]))
            fin = max(fin, y + 1)
        trozo = img.crop((0, y, img.width, fin))
        w, h = trozo.width * pt_px, trozo.height * pt_px
        n += 1
        c.drawImage(guardar(trozo, f"p{n}.png"), (pw - w) / 2, ph - h, width=w, height=h)
        c.showPage()
        y = fin
    c.save()


# ----------------------------------------------------------------- varias páginas -> 1 PDF
COMPRESION = [   # (clave, texto, resolución máx. en el PDF, calidad JPEG); None = sin pérdida
    ("sin", "Sin comprimir (máxima calidad)", None, None),
    ("normal", "Normal: buena calidad (recomendado)", 150, 75),
    ("alta", "Alta: archivo pequeño", 120, 60),
    ("maxima", "Máxima: lo más pequeño posible", 96, 45),
]


def cargar_imagenes(ruta):
    """Abre un archivo de imagen y devuelve [(imagen RGB, dpi)]. Un TIFF de varias páginas devuelve todas;
    las fotos del celular se enderezan según su orientación (EXIF)."""
    from PIL import ImageOps, ImageSequence
    res = []
    with Image.open(ruta) as src:
        for fr in ImageSequence.Iterator(src):
            d = fr.info.get("dpi") or src.info.get("dpi") or (DPI, DPI)
            d = float(d[0]) if d and d[0] and d[0] >= 30 else float(DPI)   # sin dpi o absurdo: se asume DPI
            res.append((ImageOps.exif_transpose(fr).convert("RGB"), d))
    return res


def generar_pdf_paginas(paginas, ruta, hoja=letter, dpi_max=None, calidad=None, gris=False, progreso=None):
    """Una imagen por hoja, todas en un solo PDF.
    paginas: [(imagen PIL, dpi)].
    hoja: (ancho, alto) en puntos -> la imagen va a tamaño real, centrada arriba, y se reduce solo si no cabe;
          None -> cada hoja mide lo mismo que su imagen.
    dpi_max/calidad: compresión (baja la resolución a dpi_max y guarda en JPEG con esa calidad);
          None -> PNG sin pérdida.
    gris: escala de grises (más ligero). progreso(n, total) se llama tras cada hoja."""
    c = canvas.Canvas(ruta, pagesize=hoja or letter, pageCompression=1)
    c.setTitle("Documento escaneado")
    tmpdir = tempfile.mkdtemp()
    try:
        for n, (im, d) in enumerate(paginas, 1):
            wpt, hpt = im.width / d * 72.0, im.height / d * 72.0      # tamaño real en puntos
            if hoja:
                pw, ph = hoja
                s = min(1.0, pw / wpt, ph / hpt)
            else:
                pw, ph, s = wpt, hpt, 1.0
            w, h = wpt * s, hpt * s
            im = im.convert("L" if gris else "RGB")
            if dpi_max:
                ef = im.width / (w / 72.0)                            # resolución con la que queda impresa
                if ef > dpi_max * 1.05:
                    f = dpi_max / ef
                    im = im.resize((max(1, round(im.width * f)), max(1, round(im.height * f))), Image.LANCZOS)
            if calidad:
                p = os.path.join(tmpdir, f"p{n}.jpg")                 # el JPEG se incrusta tal cual en el PDF
                im.save(p, "JPEG", quality=calidad, optimize=True)
            else:
                p = os.path.join(tmpdir, f"p{n}.png")
                im.save(p, "PNG")
            c.setPageSize((pw, ph))
            c.drawImage(p, (pw - w) / 2, ph - h, width=w, height=h)
            c.showPage()
            if progreso:
                progreso(n, len(paginas))
        c.save()
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def tamano_legible(b):
    return f"{b / 1024 / 1024:.1f} MB" if b >= 1024 * 1024 else f"{max(1, round(b / 1024))} KB"


# ----------------------------------------------------------------- comprimir un PDF existente
def _niveles_pdf(pasos=30):
    """(resolución máx. de las imágenes, calidad JPEG, calidad que se le muestra al usuario), de más suave
    a más fuerte: la calidad baja de 95 a 20 y la resolución de 300 a 50 dpi, en pasos pequeños para
    acercarse lo más posible al porcentaje pedido."""
    res = []
    for k in range(pasos + 1):
        q = round(95 - 75 * k / pasos)
        dpi = round(300 * (50 / 300) ** (k / pasos))
        txt = "alta" if q >= 70 else "media" if q >= 50 else "baja" if q >= 32 else "muy baja"
        res.append((dpi, q, txt))
    return res


NIVELES_PDF = _niveles_pdf()
HISTORIAL = os.path.join(CARPETA, "historial_compresion.json")
VERSIONES = os.path.join(CARPETA, "versiones_pdf")   # copia de cada versión comprimida (para el historial)


def info_pdf(ruta):
    """(páginas, imágenes) de un PDF. Error claro si está protegido o dañado."""
    import pymupdf
    try:
        doc = pymupdf.open(ruta)
    except Exception as e:
        raise RuntimeError(f"No se pudo abrir el PDF (¿está dañado?): {e}")
    try:
        if doc.needs_pass:
            raise RuntimeError("El PDF está protegido con contraseña.")
        imagenes = len({x[0] for pag in doc for x in pag.get_images(full=True)})
        return doc.page_count, imagenes
    finally:
        doc.close()


def _imagenes_reducibles(doc):
    """{xref: (página, resolución máx. a la que se ve)} de las imágenes que vale la pena recomprimir:
    sin transparencia, de más de 1 bit (las de 1 bit ya son muy compactas) y no pequeñas (logos)."""
    usos = {}
    for pag in doc:
        for xref, smask, w, h, bpc, *_ in pag.get_images(full=True):
            if smask or bpc == 1 or min(w, h) < 64:
                continue
            for r in pag.get_image_rects(xref):
                lado = max(r.width, r.height)
                if lado > 0:
                    d = max(w, h) / (lado / 72.0)
                    if xref not in usos or d > usos[xref][1]:
                        usos[xref] = (pag.number, d)
    return usos


def _recomprimir_imagenes(doc, dpi_max, calidad):
    """Baja la resolución de las imágenes del PDF a dpi_max (según el tamaño al que se ven en la hoja) y las
    guarda en JPEG con esa calidad. No reemplaza las que crecerían."""
    import io
    import pymupdf
    for xref, (num, d) in _imagenes_reducibles(doc).items():
        try:
            pix = pymupdf.Pixmap(doc, xref)
            if pix.alpha:
                pix = pymupdf.Pixmap(pix, 0)
            if pix.n not in (1, 3):                    # CMYK u otros -> RGB
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            im = Image.frombytes("L" if pix.n == 1 else "RGB", (pix.width, pix.height), pix.samples)
        except Exception:
            continue                                   # espacio de color raro: se deja como está
        if d > dpi_max * 1.05:
            f = dpi_max / d
            im = im.resize((max(1, round(im.width * f)), max(1, round(im.height * f))), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=calidad, optimize=True)
        if buf.tell() < len(doc.xref_stream_raw(xref)):
            doc[num].replace_image(xref, stream=buf.getvalue())     # aplica en todas las páginas que la usan


def comprimir_pdf(datos, porcentaje, progreso=None):
    """Reduce un PDF (bytes) buscando que pese 'porcentaje' % del original. Recomprime sus imágenes;
    el texto y los dibujos no se tocan. Primero prueba solo limpiar el archivo (sin pérdida); si no
    alcanza, busca el nivel de NIVELES_PDF más suave que lo logre (si ninguno, el más fuerte).
    Devuelve (bytes del PDF nuevo, nivel): nivel -1 = sin pérdida, si no índice en NIVELES_PDF.
    Si no se pudo reducir nada devuelve (None, None). progreso(n) se llama antes de cada prueba."""
    import pymupdf
    objetivo = len(datos) * porcentaje / 100.0
    hechos = {}
    con_imagenes = []

    def probar(i):                                     # i = -1: solo limpieza, sin tocar imágenes
        if i not in hechos:
            if progreso:
                progreso(len(hechos) + 1)
            doc = pymupdf.open(stream=datos, filetype="pdf")
            try:
                if doc.needs_pass:
                    raise RuntimeError("El PDF está protegido con contraseña.")
                if i >= 0:
                    _recomprimir_imagenes(doc, *NIVELES_PDF[i][:2])
                else:
                    con_imagenes.append(bool(_imagenes_reducibles(doc)))
                hechos[i] = doc.tobytes(garbage=4, deflate=True, use_objstms=1, clean=True)
            finally:
                doc.close()
        return hechos[i]

    elegido = -1
    if len(probar(-1)) > objetivo and con_imagenes[0]:  # búsqueda binaria del nivel más suave que alcanza
        lo, hi = 0, len(NIVELES_PDF) - 1
        while lo < hi:
            mid = (lo + hi) // 2
            if len(probar(mid)) <= objetivo:
                hi = mid
            else:
                lo = mid + 1
        elegido = hi
        if len(probar(hi)) > objetivo:                 # ninguno alcanza: el que más redujo
            elegido = min(hechos, key=lambda i: len(hechos[i]))
    if len(hechos[elegido]) >= len(datos):
        return None, None
    return hechos[elegido], elegido


def ruta_comprimido(origen, porcentaje):
    base = os.path.splitext(origen)[0]
    return f"{base}_comprimido_{porcentaje}.pdf"


def guardar_comprimido(origen, porcentaje, datos):
    """Guarda el PDF comprimido junto al original (el original no se toca). Si esa carpeta no deja
    escribir, lo guarda en CARPETA. Devuelve la ruta final."""
    destino = ruta_comprimido(origen, porcentaje)
    for intento in (destino, os.path.join(CARPETA, os.path.basename(destino))):
        tmp = intento + ".tmp"
        try:
            os.makedirs(os.path.dirname(intento), exist_ok=True)
            with open(tmp, "wb") as f:
                f.write(datos)
        except PermissionError:
            continue                                   # carpeta sin permiso de escritura: probar la otra
        try:
            os.replace(tmp, intento)
        except PermissionError:
            try:
                os.remove(tmp)
            except OSError:
                pass
            raise RuntimeError(f"No se pudo guardar «{os.path.basename(intento)}» porque está abierto en otro "
                               "programa. Ciérralo e intenta de nuevo.")
        return intento
    raise RuntimeError("No se pudo guardar el PDF comprimido: no hay permiso para escribir en la carpeta.")


def cargar_historial():
    try:
        with open(HISTORIAL, encoding="utf8") as f:
            h = json.load(f)
        return h if isinstance(h, list) else []
    except Exception:
        return []


def _guardar_historial(h):
    os.makedirs(os.path.dirname(HISTORIAL), exist_ok=True)
    with open(HISTORIAL, "w", encoding="utf8") as f:
        json.dump(h, f, ensure_ascii=False, indent=2)


def _borrar_copias(entradas):
    """Borra las copias guardadas para el historial (solo archivos dentro de VERSIONES)."""
    for x in entradas:
        c = x.get("copia")
        if c and _clave_ruta(os.path.dirname(c)) == _clave_ruta(VERSIONES):
            try:
                os.remove(c)
            except OSError:
                pass


def agregar_historial(entrada):
    h = cargar_historial() + [entrada]
    _borrar_copias(h[:-1000])                          # máximo 1000 versiones; las más viejas se descartan
    _guardar_historial(h[-1000:])


def quitar_de_historial(entrada):
    """Quita una versión del historial y borra su copia guardada (el archivo junto al original no se toca)."""
    _guardar_historial([x for x in cargar_historial() if x != entrada])
    _borrar_copias([entrada])


def archivar_version(datos):
    """Guarda una copia de una versión comprimida, para poder volver a guardarla desde el historial
    aunque el archivo junto al original se mueva o se borre. Devuelve su ruta."""
    import uuid
    os.makedirs(VERSIONES, exist_ok=True)
    ruta = os.path.join(VERSIONES, datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid.uuid4().hex[:8] + ".pdf")
    with open(ruta, "wb") as f:
        f.write(datos)
    return ruta


def version_disponible(h):
    """De dónde se puede abrir o guardar una versión del historial: su copia guardada o, si no, el archivo."""
    for r in (h.get("copia"), h.get("salida")):
        if r and os.path.exists(r):
            return r
    return None


def _clave_ruta(ruta):
    return os.path.normcase(os.path.abspath(ruta))


def historial_de(ruta):
    """Compresiones del mismo documento: las hechas a este archivo, a su original y a sus copias comprimidas."""
    hist = cargar_historial()
    padre = {_clave_ruta(h["salida"]): _clave_ruta(h["origen"]) for h in hist}

    def raiz(c):
        vistos = set()
        while c in padre and c not in vistos:
            vistos.add(c)
            c = padre[c]
        return c

    r = raiz(_clave_ruta(ruta))
    return [h for h in hist if raiz(_clave_ruta(h["origen"])) == r]


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
        im = Image.new("RGBA", (300, 300), (0, 0, 0, 0))     # lienzo holgado: el emoji no se corta
        ImageDraw.Draw(im).text((60, 60), emoji, font=fuente, embedded_color=True)
        im = im.crop(im.getbbox() or (0, 0, 300, 300))
        lado = max(im.size)
        cuadro = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
        cuadro.paste(im, ((lado - im.width) // 2, (lado - im.height) // 2))
        ph = ImageTk.PhotoImage(cuadro.resize((px, px), Image.LANCZOS))
    except Exception:
        ph = None
    _ICO_CACHE[clave] = ph
    return ph


FUENTE_ICONOS = next((f for f in (r"C:\Windows\Fonts\segmdl2.ttf", r"C:\Windows\Fonts\SegoeIcons.ttf")
                      if os.path.exists(f)), None)


def icono_plano(glifo, color, px=36, emoji=None):
    """Icono plano: cuadro de color con esquinas redondeadas y un símbolo blanco (de la fuente de iconos
    de Windows). Si esa fuente no existe, se usa el emoji."""
    from PIL import ImageFont
    clave = ("plano", glifo, color, px)
    if clave in _ICO_CACHE:
        return _ICO_CACHE[clave]
    ph = None
    if FUENTE_ICONOS and glifo:
        try:
            n = px * 4                                          # supermuestreo: bordes suaves
            im = Image.new("RGBA", (n, n), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.rounded_rectangle((0, 0, n - 1, n - 1), radius=int(n * 0.2), fill=color)
            d.text((n / 2, n / 2), glifo, font=ImageFont.truetype(FUENTE_ICONOS, int(n * 0.56)),
                   fill="white", anchor="mm")
            ph = ImageTk.PhotoImage(im.resize((px, px), Image.LANCZOS))
        except Exception:
            ph = None
    if ph is None and emoji:
        ph = icono(emoji, px)
    _ICO_CACHE[clave] = ph
    return ph


def fondo_redondeado(w, h, radio, relleno, borde, fondo, grosor=1):
    """Rectángulo con esquinas redondeadas y bordes suaves (tarjetas y pastillas), sobre el color 'fondo'."""
    clave = ("rr", w, h, radio, relleno, borde, fondo, grosor)
    if clave in _ICO_CACHE:
        return _ICO_CACHE[clave]
    S = 4
    im = Image.new("RGB", (w * S, h * S), fondo)
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((0, 0, w * S - 1, h * S - 1), radius=radio * S, fill=borde)
    g = grosor * S
    d.rounded_rectangle((g, g, w * S - 1 - g, h * S - 1 - g), radius=max(0, radio * S - g), fill=relleno)
    ph = ImageTk.PhotoImage(im.resize((w, h), Image.LANCZOS))
    _ICO_CACHE[clave] = ph
    return ph


def logo_app(px=28):
    """El icono de la aplicación (icono.ico) en tamaño px; si no está, el emoji de credencial."""
    clave = ("logo", px)
    if clave not in _ICO_CACHE:
        try:
            with Image.open(ICONO_APP) as im:
                _ICO_CACHE[clave] = ImageTk.PhotoImage(im.convert("RGBA").resize((px, px), Image.LANCZOS))
        except Exception:
            _ICO_CACHE[clave] = icono("🪪", px)
    return _ICO_CACHE[clave]


# ----------------------------------------------------------------- interfaz
# Una sola ventana: la pantalla de inicio muestra el título y las funciones; al elegir una, se oculta
# el inicio y se muestra esa función en la misma ventana, con un botón para regresar.
TIPOS_IMAGEN = "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.gif"


def _boton(padre, emoji, texto, cmd, px=26):
    return ttk.Button(padre, text=" " + texto, image=icono(emoji, px),
                      compound="left", command=cmd, style="Icono.TButton")


def _etiqueta(padre, emoji, texto, px=24, **kw):
    return ttk.Label(padre, text=" " + texto, image=icono(emoji, px), compound="left", **kw)


def pedir_imagen(padre):
    """Abre una imagen elegida por el usuario (con su dpi en .dpi_real), o None si cancela o falla."""
    ruta = filedialog.askopenfilename(parent=padre, filetypes=[("Imágenes", TIPOS_IMAGEN),
                                                               ("Todos los archivos", "*.*")])
    if not ruta:
        return None
    try:
        im, d = cargar_imagenes(ruta)[0]
    except Exception as e:
        messagebox.showerror("❌ Imagen", f"❌ No se pudo abrir la imagen:\n{e}", parent=padre)
        return None
    im.dpi_real = d
    return im


class Vista(ttk.Frame):
    """Una pantalla de la ventana principal (el inicio o una función)."""
    EMOJI, TITULO, RESUMEN, AYUDA = "", "", "", ""
    GLIFO, COLOR, CATEGORIA, NUEVO = "", "#888888", "", False   # tarjeta del inicio
    USA_ESCANER = True                                 # muestra la barra de escáner/papel arriba

    def __init__(self, app):
        super().__init__(app.cuerpo)
        self.app = app
        self.ico_estado = self.estado = None

    def al_mostrar(self):
        """Se llama cada vez que la pantalla se vuelve visible."""

    def _luego(self, fn):
        """Ejecuta fn en el hilo de la interfaz (se llama desde un hilo de trabajo)."""
        try:
            self.app.after(0, fn)
        except Exception:
            pass                                       # la ventana ya se cerró

    def barra_estado(self, padre, ancho=520):
        self.ico_estado = ttk.Label(padre)
        self.ico_estado.pack(side="left")
        self.estado = ttk.Label(padre, text="", wraplength=ancho, justify="left")
        self.estado.pack(side="left", padx=6)

    def msg(self, t):
        """El primer símbolo del mensaje (✅ ⚠️ ❌ ⏳ ℹ️ 🧹) se muestra como icono grande."""
        t = t.strip()
        e_ini = ""
        for e in ("✅", "⚠️", "❌", "⏳", "ℹ️", "🧹"):
            if t.startswith(e):
                e_ini, t = e, t[len(e):].strip()
                break
        t = t.replace("🔁 ", "").replace("🪪 ", "")
        self.ico_estado.config(image=icono(e_ini, 30) if e_ini else "")
        self.estado.config(text=t)
        self.update_idletasks()


FONDO_INICIO, TINTA, TINTA_SUAVE, ACENTO = "#f3f3f6", "#2b2b33", "#5f5f6b", "#e0403a"


class Inicio(Vista):
    """Pantalla de inicio: título, filtros por categoría y una tarjeta por cada función."""
    TITULO = "INEScan 2x"
    BORDE, BORDE_ACTIVO = "#e1e1e7", TINTA
    TW, TH, COLUMNAS = 208, 176, 4                     # tamaño de cada tarjeta y tarjetas por fila

    def __init__(self, app, funciones):
        super().__init__(app)
        import tkinter.font as tkfont
        F = FONDO_INICIO
        pag = tk.Frame(self, bg=F)
        pag.pack(fill="both", expand=True)

        pie = tk.Frame(pag, bg=F)                       # estado del escáner, abajo
        pie.pack(side="bottom", fill="x", padx=16, pady=(0, 12))
        self.ico_estado = tk.Label(pie, bg=F)
        self.ico_estado.pack(side="left")
        self.estado = tk.Label(pie, bg=F, fg=TINTA, font=("Segoe UI", 9), wraplength=820, justify="left")
        self.estado.pack(side="left", padx=6)

        tk.Label(pag, text="Herramientas para tus credenciales y documentos", bg=F, fg=TINTA,
                 font=("Segoe UI", -30, "bold")).pack(pady=(34, 6))
        tk.Label(pag, text="Escanea tu INE y tus documentos, júntalos en un solo PDF y hazlos más ligeros. "
                           "Todo se hace en tu computadora: tus archivos no se suben a internet.",
                 bg=F, fg=TINTA_SUAVE, font=("Segoe UI", -14), wraplength=620, justify="center").pack()

        # filtros por categoría (solo las que tienen funciones)
        cats = ["Todos"]
        for _, cls in funciones:
            if cls.CATEGORIA not in cats:
                cats.append(cls.CATEGORIA)
        fila = tk.Frame(pag, bg=F)
        fila.pack(pady=(22, 16))
        self.fuente_pastilla = tkfont.Font(family="Segoe UI", size=-12, weight="bold")
        self.pastillas = {}
        for cat in cats:
            self.pastillas[cat] = self._pastilla(fila, cat)
            self.pastillas[cat].pack(side="left", padx=4)

        self.rejilla = tk.Frame(pag, bg=F)
        self.rejilla.pack()
        self.tarjetas = []                             # [(categoría, canvas, pintar)]
        for clave, cls in funciones:
            self.tarjetas.append((cls.CATEGORIA,) + self._tarjeta(self.rejilla, cls, lambda c=clave: app.ir(c)))
        self.categoria = None
        self.filtrar("Todos")

    def al_mostrar(self):
        for _, _, pintar in self.tarjetas:             # al regresar, ninguna tarjeta queda resaltada
            pintar(False)

    # --- filtros
    def _pastilla(self, padre, texto):
        w, h = self.fuente_pastilla.measure(texto) + 28, 28
        c = tk.Canvas(padre, width=w, height=h, bg=FONDO_INICIO, highlightthickness=0, cursor="hand2")
        c.fondo = c.create_image(0, 0, anchor="nw")
        c.texto = c.create_text(w / 2, h / 2, text=texto, font=self.fuente_pastilla)
        c.medidas = (w, h)
        c.bind("<Button-1>", lambda e: self.filtrar(texto))
        c.bind("<Enter>", lambda e: self._pintar_pastilla(texto, True))
        c.bind("<Leave>", lambda e: self._pintar_pastilla(texto))
        return c

    def _pintar_pastilla(self, cat, encima=False):
        c = self.pastillas[cat]
        w, h = c.medidas
        activa = cat == self.categoria
        relleno = TINTA if activa else "#ffffff"
        borde = TINTA if activa or encima else "#d3d3da"
        c.itemconfig(c.fondo, image=fondo_redondeado(w, h, h // 2, relleno, borde, FONDO_INICIO))
        c.itemconfig(c.texto, fill="#ffffff" if activa else TINTA)

    def filtrar(self, cat):
        self.categoria = cat
        for c in self.pastillas:
            self._pintar_pastilla(c)
        for _, t, _ in self.tarjetas:
            t.grid_forget()
        visibles = [t for c, t, _ in self.tarjetas if cat == "Todos" or c == cat]
        for i, t in enumerate(visibles):
            t.grid(row=i // self.COLUMNAS, column=i % self.COLUMNAS, padx=6, pady=6)

    # --- tarjetas
    def _tarjeta(self, padre, cls, cmd):
        """Tarjeta clicable: icono, nombre de la función y una descripción corta. Devuelve (canvas, pintar)."""
        W, H = self.TW, self.TH
        c = tk.Canvas(padre, width=W, height=H, bg=FONDO_INICIO, highlightthickness=0, cursor="hand2",
                      takefocus=1)
        fondo = c.create_image(0, 0, anchor="nw")
        c.create_image(20, 20, anchor="nw", image=icono_plano(cls.GLIFO, cls.COLOR, 38, cls.EMOJI))
        if cls.NUEVO:
            c.create_image(W - 12, 12, anchor="ne", image=fondo_redondeado(52, 18, 4, "#e4efff", "#e4efff", "#ffffff"))
            c.create_text(W - 38, 21, text="¡Nuevo!", font=("Segoe UI", -10, "bold"), fill="#2f6fd1")
        c.create_text(20, 74, anchor="nw", text=cls.TITULO, font=("Segoe UI", -15, "bold"), fill=TINTA)
        c.create_text(20, 100, anchor="nw", text=cls.RESUMEN, font=("Segoe UI", -11), fill=TINTA_SUAVE,
                      width=W - 40)

        def pintar(activa):
            c.itemconfig(fondo, image=fondo_redondeado(W, H, 10, "#ffffff",
                                                       self.BORDE_ACTIVO if activa else self.BORDE, FONDO_INICIO))
        pintar(False)
        c.bind("<Enter>", lambda e: pintar(True))
        c.bind("<Leave>", lambda e: pintar(False))
        c.bind("<FocusIn>", lambda e: pintar(True))
        c.bind("<FocusOut>", lambda e: pintar(False))
        for ev in ("<Button-1>", "<Return>", "<space>"):
            c.bind(ev, lambda e: cmd())
        return c, pintar


class Credencial(Vista):
    """INE / credencial: frente y reverso -> PDF al doble o tamaño normal."""
    EMOJI, TITULO = "🪪", "Credencial (INE)"
    GLIFO, COLOR, CATEGORIA = "", "#2f7be5", "Escanear"
    RESUMEN = ("Escanea el frente y el reverso de la INE y obtén un PDF listo para imprimir, "
               "al doble o a tamaño normal.")
    AYUDA = ("1) Pon la credencial en el cristal, en la esquina superior, y escanea el FRENTE.  "
             "2) Voltéala y escanea el REVERSO.  3) Elige el PDF al doble o tamaño normal.")

    def __init__(self, app):
        super().__init__(app)
        self.frente = None
        self.reverso = None

        mid = ttk.Frame(self, padding=10)
        mid.pack(fill="both", expand=True)
        self.lbl = {}
        for i, (clave, emoji, texto) in enumerate((("frente", "🪪", "FRENTE"), ("reverso", "🔁", "REVERSO"))):
            f = ttk.LabelFrame(mid, labelwidget=_etiqueta(mid, emoji, texto, 28), padding=6)
            f.grid(row=0, column=i, padx=8, sticky="nsew")
            mid.columnconfigure(i, weight=1, uniform="col")
            l = ttk.Label(f, text="(sin escanear)", image=icono("🖼️", 72), compound="top", anchor="center")
            l.pack(fill="both", expand=True)
            self.lbl[clave] = l
            _boton(f, "📷", "Escanear " + texto.lower(),
                   lambda c=clave: self.escanear(c)).pack(pady=4, fill="x")
            _boton(f, "📂", "Abrir imagen…",
                   lambda c=clave: self.abrir(c)).pack(fill="x")
        mid.rowconfigure(0, weight=1)

        bot = ttk.Frame(self, padding=10)
        bot.pack(fill="x")
        self.barra_estado(bot, 360)
        _boton(bot, "📄", "PDF tamaño normal", lambda: self.pdf(1)).pack(side="right")
        _boton(bot, "🔍", "PDF al doble (2x)", lambda: self.pdf(2)).pack(side="right", padx=6)
        _boton(bot, "🧹", "Limpiar", self.limpiar).pack(side="right", padx=6)
        self.msg("ℹ️ Pon la credencial en el cristal, esquina superior, y escanea el FRENTE.")

    def mostrar(self, clave, im):
        mini = im.copy()
        mini.thumbnail((380, 250))
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
        self.app.escanear_async(lambda im: self.procesar(clave, im), self.msg, f"⏳ Escaneando {clave}… espera.")

    def abrir(self, clave):
        im = pedir_imagen(self)
        if im is not None:
            self.procesar(clave, im)

    def limpiar(self):
        self.frente = self.reverso = None
        for l in self.lbl.values():
            l.config(image=icono("🖼️", 72), text="(sin escanear)", compound="top")
            l.image = None
        self.msg("🧹 Listo. Escanea el FRENTE.")

    def pdf(self, factor=2):
        if self.frente is None or self.reverso is None:
            messagebox.showinfo("ℹ️ Falta algo", "ℹ️ Escanea primero 🪪 frente y 🔁 reverso.", parent=self)
            return
        os.makedirs(CARPETA, exist_ok=True)
        ruta = os.path.join(CARPETA, datetime.now().strftime("credencial_%Y%m%d_%H%M%S")
                            + ("_normal.pdf" if factor == 1 else "_2x.pdf"))
        try:
            generar_pdf(self.frente, self.reverso, ruta, self.app.pagina(), factor=factor)
        except Exception as e:
            messagebox.showerror("❌ PDF", "❌ No se pudo crear el PDF:\n" + str(e), parent=self)
            return
        self.msg(f"✅ PDF guardado: {ruta}")
        os.startfile(ruta)


class DocOficio(Vista):
    """Documento más largo que el cristal: se escanea en 2 partes, se unen solas y se elige cómo imprimir."""
    EMOJI, TITULO = "📑", "Documento oficio"
    GLIFO, COLOR, CATEGORIA = "", "#eb6e3d", "Escanear"
    RESUMEN = ("Escanea en 2 partes documentos más largos que el cristal; se unen, enderezan y "
               "recortan solos.")
    AYUDA = ("1) Escanea la parte de ARRIBA.  2) Desliza el documento y escanea lo que falta, "
             "dejando ~2 cm repetidos.  3) Elige cómo imprimir.")
    FORMAS = [("oficio", "Oficio, tamaño real (21.6 x 34 cm)"),
              ("carta", "Ajustar a 1 hoja carta/A4 (reducido)"),
              ("dividir", "Dividir en hojas carta/A4 (tamaño real)")]

    def __init__(self, app):
        super().__init__(app)
        self.partes = {"1": None, "2": None}
        self.unido = None

        mid = ttk.Frame(self, padding=10)
        mid.pack(fill="both", expand=True)
        self.lbl = {}
        for i, (clave, emoji, titulo) in enumerate((("1", "⬆️", "PARTE 1 (arriba)"),
                                                    ("2", "⬇️", "PARTE 2 (lo faltante)"),
                                                    ("res", "🧩", "RESULTADO UNIDO"))):
            f = ttk.LabelFrame(mid, labelwidget=_etiqueta(mid, emoji, titulo, 24), padding=6)
            f.grid(row=0, column=i, padx=6, sticky="nsew")
            mid.columnconfigure(i, weight=1, uniform="col")
            l = ttk.Label(f, text="(vacío)", image=icono("🖼️", 72), compound="top", anchor="center")
            l.pack(fill="both", expand=True)
            self.lbl[clave] = l
            if clave != "res":
                _boton(f, "📷", "Escanear", lambda c=clave: self.escanear(c)).pack(pady=4, fill="x")
                _boton(f, "📂", "Abrir imagen…", lambda c=clave: self.abrir(c)).pack(fill="x")
        mid.rowconfigure(0, weight=1)

        bot = ttk.Frame(self, padding=(10, 6, 10, 4))
        bot.pack(fill="x")
        _etiqueta(bot, "🖨️", "Forma de impresión:").pack(side="left")
        self.forma = ttk.Combobox(bot, state="readonly", width=38, values=[t for _, t in self.FORMAS])
        self.forma.current(0)
        self.forma.pack(side="left", padx=6)
        _boton(bot, "🧹", "Limpiar", self.limpiar).pack(side="right")
        _boton(bot, "📄", "Generar PDF", self.pdf).pack(side="right", padx=6)

        est = ttk.Frame(self, padding=(12, 2, 12, 10))
        est.pack(fill="x")
        self.barra_estado(est, 860)
        self.msg("ℹ️ Escanea la parte 1 (la de ARRIBA del documento).")

    def mostrar(self, clave, im):
        mini = im.copy()
        mini.thumbnail((260, 300))
        ph = ImageTk.PhotoImage(mini)
        self.lbl[clave].config(image=ph, text="", compound="center")
        self.lbl[clave].image = ph

    def procesar(self, clave, im):
        self.partes[clave] = im
        self.mostrar(clave, im)
        if self.partes["1"] is not None and self.partes["2"] is not None:
            self.unir()
        elif clave == "1":
            self.msg("ℹ️ Parte 1 lista. Desliza el documento y escanea la parte 2 (con ~2 cm de traslape).")

    def unir(self):
        try:
            img, info = unir_partes(self.partes["1"], self.partes["2"])
        except Exception as e:
            messagebox.showerror("Unir", f"No se pudo unir:\n{e}", parent=self)
            return
        self.unido = img
        self.mostrar("res", img)
        if info is None:
            self.msg("⚠️ No encontré dónde continúa el documento; las partes se pusieron una tras otra. "
                     "Reescanea la parte 2 dejando ~2 cm repetidos de la parte 1 (que tengan texto o dibujo).")
        else:
            alto_cm = img.height / getattr(img, "dpi_real", DPI) * 2.54
            self.msg(f"✅ Unido ({info}). Alto total: {alto_cm:.1f} cm. "
                     "Elige la forma de impresión y genera el PDF.")

    def escanear(self, clave):
        self.app.escanear_async(lambda im: self.procesar(clave, im), self.msg,
                                f"⏳ Escaneando parte {clave}… espera.")

    def abrir(self, clave):
        im = pedir_imagen(self)
        if im is not None:
            self.procesar(clave, im)

    def limpiar(self):
        self.partes = {"1": None, "2": None}
        self.unido = None
        for l in self.lbl.values():
            l.config(image=icono("🖼️", 72), text="(vacío)", compound="top")
            l.image = None
        self.msg("🧹 Listo. Escanea la parte 1 (la de ARRIBA del documento).")

    def pdf(self):
        if self.unido is None:
            messagebox.showinfo("Falta algo", "Escanea primero la parte 1 y la parte 2.", parent=self)
            return
        modo = self.FORMAS[self.forma.current()][0]
        os.makedirs(CARPETA, exist_ok=True)
        ruta = os.path.join(CARPETA, datetime.now().strftime("documento_%Y%m%d_%H%M%S") + f"_{modo}.pdf")
        try:
            generar_pdf_documento(self.unido, ruta, modo, self.app.pagina())
        except Exception as e:
            messagebox.showerror("PDF", "No se pudo crear el PDF:\n" + str(e), parent=self)
            return
        self.msg(f"✅ PDF guardado: {ruta}")
        os.startfile(ruta)


class VariasPaginas(Vista):
    """Escanea (o agrega) varias hojas o imágenes y las guarda en un solo PDF, con opción de comprimirlo."""
    EMOJI, TITULO = "🗂️", "Varias páginas"
    GLIFO, COLOR, CATEGORIA, NUEVO = "", "#9b51e0", "Escanear", True
    RESUMEN = ("Escanea varias hojas o junta imágenes en un solo PDF. Ordénalas, gíralas y elige "
               "la compresión.")
    AYUDA = ("1) Pon una hoja en el cristal y presiona Escanear página; repite con cada hoja "
             "(o agrega imágenes).  2) Ordénalas y gíralas si hace falta.  3) Elige la compresión y genera el PDF.")

    def __init__(self, app):
        super().__init__(app)
        self.paginas = []          # [{"img": PIL RGB, "dpi": float, "nombre": str, "mini": PhotoImage o None}]
        self.generando = False
        cfg = cargar_config()

        mid = ttk.Frame(self, padding=10)
        mid.pack(fill="both", expand=True)
        izq = ttk.LabelFrame(mid, labelwidget=_etiqueta(mid, "📚", "PÁGINAS", 24), padding=6)
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        der = ttk.LabelFrame(mid, labelwidget=_etiqueta(mid, "🖼️", "VISTA PREVIA", 24), padding=6)
        der.grid(row=0, column=1, sticky="nsew")
        mid.columnconfigure(1, weight=1)
        mid.rowconfigure(0, weight=1)

        _boton(izq, "📷", "Escanear página", self.escanear).pack(fill="x")
        _boton(izq, "📂", "Agregar imágenes…", self.agregar).pack(fill="x", pady=(4, 6))
        caja = ttk.Frame(izq)
        caja.pack(fill="both", expand=True)
        self.lista = tk.Listbox(caja, width=34, height=8, activestyle="none", exportselection=False,
                                font=("Segoe UI", 10))
        barra = ttk.Scrollbar(caja, orient="vertical", command=self.lista.yview)
        self.lista.config(yscrollcommand=barra.set)
        self.lista.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.lista.bind("<<ListboxSelect>>", lambda e: self.ver())
        self.lista.bind("<Delete>", lambda e: self.quitar())
        acc = ttk.Frame(izq)
        acc.pack(fill="x", pady=(6, 0))
        for i, (emoji, texto, cmd) in enumerate((("⬆️", "Subir", lambda: self.mover(-1)),
                                                 ("⬇️", "Bajar", lambda: self.mover(1)),
                                                 ("🔃", "Girar", self.girar),
                                                 ("🗑️", "Quitar", self.quitar))):
            _boton(acc, emoji, texto, cmd, px=20).grid(row=i // 2, column=i % 2, sticky="ew", padx=2, pady=2)
            acc.columnconfigure(i % 2, weight=1)

        self.vista = ttk.Label(der, anchor="center")
        self.vista.pack(fill="both", expand=True)
        self.info = ttk.Label(der, text="", foreground="#555", anchor="center")
        self.info.pack(fill="x")

        opc = ttk.Frame(self, padding=(10, 0, 10, 4))
        opc.pack(fill="x")
        _etiqueta(opc, "🗜️", "Compresión:").pack(side="left")
        self.comp = ttk.Combobox(opc, state="readonly", width=34, values=[t for _, t, _, _ in COMPRESION])
        claves = [k for k, _, _, _ in COMPRESION]
        self.comp.current(claves.index(cfg.get("compresion")) if cfg.get("compresion") in claves else 1)
        self.comp.pack(side="left", padx=(4, 14))
        self.gris = tk.BooleanVar(value=bool(cfg.get("gris", False)))
        ttk.Checkbutton(opc, text="Escala de grises (más ligero)", variable=self.gris).pack(side="left")
        self.hoja_imagen = tk.BooleanVar(value=False)
        ttk.Checkbutton(opc, text="Hoja del tamaño de cada imagen", variable=self.hoja_imagen).pack(side="left", padx=14)

        bot = ttk.Frame(self, padding=(10, 4, 10, 10))
        bot.pack(fill="x")
        self.barra_estado(bot, 540)
        _boton(bot, "🧹", "Limpiar", self.limpiar).pack(side="right")
        self.btn_pdf = _boton(bot, "📄", "Generar PDF", self.pdf)
        self.btn_pdf.pack(side="right", padx=6)

        self.ver()
        self.msg("ℹ️ Pon la primera hoja en el cristal y presiona Escanear página.")

    # --- lista de páginas
    def actual(self):
        s = self.lista.curselection()
        return s[0] if s else None

    def refrescar(self, sel=None):
        self.lista.delete(0, "end")
        for n, p in enumerate(self.paginas, 1):
            self.lista.insert("end", f"{n:>3}.  {p['nombre']}")
        if self.paginas:
            sel = len(self.paginas) - 1 if sel is None else max(0, min(sel, len(self.paginas) - 1))
            self.lista.selection_set(sel)
            self.lista.see(sel)
        self.ver()

    def ver(self):
        i = self.actual()
        if i is None:
            self.vista.config(image=icono("🖼️", 72), text="(sin páginas)", compound="top")
            self.vista.image = None
            self.info.config(text="")
            return
        p = self.paginas[i]
        if p["mini"] is None:
            mini = p["img"].copy()
            mini.thumbnail((540, 340))
            p["mini"] = ImageTk.PhotoImage(mini)
        self.vista.config(image=p["mini"], text="", compound="center")
        self.vista.image = p["mini"]
        cm = 2.54 / p["dpi"]
        self.info.config(text=f"Página {i + 1} de {len(self.paginas)}  ·  "
                              f"{p['img'].width * cm:.1f} x {p['img'].height * cm:.1f} cm")

    def mover(self, delta):
        i = self.actual()
        if i is None:
            return
        j = i + delta
        if 0 <= j < len(self.paginas):
            self.paginas[i], self.paginas[j] = self.paginas[j], self.paginas[i]
            self.refrescar(j)

    def girar(self):
        i = self.actual()
        if i is None:
            return
        p = self.paginas[i]
        p["img"] = p["img"].rotate(-90, expand=True)          # 90° a la derecha
        p["mini"] = None
        self.refrescar(i)

    def quitar(self):
        i = self.actual()
        if i is None:
            return
        del self.paginas[i]
        self.refrescar(i)
        self.msg(f"🧹 Página quitada. Quedan {len(self.paginas)}.")

    def limpiar(self):
        if self.paginas and not messagebox.askyesno(
                "Limpiar", f"¿Quitar las {len(self.paginas)} páginas?", parent=self):
            return
        self.paginas = []
        self.refrescar()
        self.msg("🧹 Listo. Pon la primera hoja en el cristal y presiona Escanear página.")

    # --- agregar páginas
    def escanear(self):
        self.app.escanear_async(self._escaneada, self.msg,
                                f"⏳ Escaneando página {len(self.paginas) + 1}… espera.")

    def _escaneada(self, im):
        self.paginas.append({"img": im, "dpi": float(getattr(im, "dpi_real", DPI)),
                             "nombre": "Escaneo " + datetime.now().strftime("%H:%M:%S"), "mini": None})
        self.refrescar()
        self.msg(f"✅ Página {len(self.paginas)} agregada. Pon la siguiente hoja y presiona Escanear página, "
                 "o genera el PDF.")

    def agregar(self):
        rutas = filedialog.askopenfilenames(
            parent=self, title="Agregar imágenes",
            filetypes=[("Imágenes", TIPOS_IMAGEN), ("Todos los archivos", "*.*")])
        if not rutas:
            return
        self.msg(f"⏳ Abriendo {len(rutas)} archivo(s)…")
        antes, malos = len(self.paginas), []
        for r in rutas:
            try:
                imgs = cargar_imagenes(r)
            except Exception as e:
                malos.append(f"{os.path.basename(r)}: {e}")
                continue
            for k, (im, d) in enumerate(imgs, 1):
                nombre = os.path.basename(r) + (f" ({k}/{len(imgs)})" if len(imgs) > 1 else "")
                self.paginas.append({"img": im, "dpi": d, "nombre": nombre, "mini": None})
        self.refrescar()
        nuevas = len(self.paginas) - antes
        if malos:
            self.msg(f"⚠️ Se agregaron {nuevas} página(s); {len(malos)} archivo(s) no se pudieron abrir.")
            messagebox.showwarning("⚠️ Imágenes", "No se pudieron abrir:\n\n" + "\n".join(malos), parent=self)
        else:
            self.msg(f"✅ Se agregaron {nuevas} página(s). Total: {len(self.paginas)}.")

    # --- PDF
    def pdf(self):
        if self.generando:
            return
        if not self.paginas:
            messagebox.showinfo("ℹ️ Falta algo", "ℹ️ Escanea o agrega al menos una página.", parent=self)
            return
        hoja = None if self.hoja_imagen.get() else self.app.pagina()       # None = tamaño de la imagen
        clave, _, dpi_max, calidad = COMPRESION[self.comp.current()]
        gris = self.gris.get()
        cfg = cargar_config()
        cfg["compresion"], cfg["gris"] = clave, gris
        guardar_config(cfg)
        os.makedirs(CARPETA, exist_ok=True)
        ruta = os.path.join(CARPETA, datetime.now().strftime("paginas_%Y%m%d_%H%M%S.pdf"))
        paginas = [(p["img"], p["dpi"]) for p in self.paginas]
        self.generando = True
        self.btn_pdf.state(["disabled"])
        self.msg(f"⏳ Generando PDF de {len(paginas)} página(s)…")

        def progreso(n, total):
            self._luego(lambda: self.msg(f"⏳ Generando PDF… página {n} de {total}"))

        def trabajo():
            try:
                generar_pdf_paginas(paginas, ruta, hoja, dpi_max, calidad, gris, progreso)
                self._luego(lambda: self._pdf_listo(ruta, len(paginas)))
            except Exception as e:
                self._luego(lambda m=str(e): self._pdf_fallido(m))
        threading.Thread(target=trabajo, daemon=True).start()

    def _pdf_listo(self, ruta, n):
        self.generando = False
        self.btn_pdf.state(["!disabled"])
        self.msg(f"✅ PDF guardado ({n} página(s), {tamano_legible(os.path.getsize(ruta))}): {ruta}")
        os.startfile(ruta)

    def _pdf_fallido(self, m):
        self.generando = False
        self.btn_pdf.state(["!disabled"])
        self.msg("❌ No se pudo crear el PDF.")
        messagebox.showerror("❌ PDF", "❌ No se pudo crear el PDF:\n" + m, parent=self)


class ComprimirPDF(Vista):
    """Reduce un PDF que ya existe al porcentaje elegido y muestra el historial de ese archivo."""
    EMOJI, TITULO = "🗜️", "Comprimir PDF"
    GLIFO, COLOR, CATEGORIA, NUEVO = "", "#5fae3e", "Optimizar PDF", True
    RESUMEN = ("Haz más ligero un PDF que ya tienes eligiendo el porcentaje. Conserva el texto y "
               "guarda el historial.")
    AYUDA = ("1) Elige un PDF.  2) Mueve la barra al porcentaje de su tamaño que quieres que pese.  "
             "3) Presiona Comprimir: se guarda una copia junto al original (el original no se toca).")
    USA_ESCANER = False
    VERDE = "#1a7f37"

    def __init__(self, app):
        super().__init__(app)
        self.ruta = None           # PDF elegido
        self.tam = 0               # su tamaño en bytes
        self.salida = None         # último PDF comprimido
        self.trabajando = False
        self.filas = {}            # fila del historial -> su entrada (versión comprimida)
        cfg = cargar_config()

        caja = ttk.Frame(self, padding=(12, 10, 12, 0))
        caja.pack(fill="both", expand=True)

        # 1) archivo
        arc = ttk.LabelFrame(caja, labelwidget=_etiqueta(caja, "📄", "ARCHIVO", 24), padding=8)
        arc.pack(fill="x")
        _boton(arc, "📂", "Elegir PDF…", self.elegir).pack(side="left")
        txt = ttk.Frame(arc)
        txt.pack(side="left", fill="x", expand=True, padx=14)
        self.lbl_nombre = ttk.Label(txt, text="(ningún PDF elegido)", font=("Segoe UI", 11, "bold"))
        self.lbl_nombre.pack(anchor="w")
        self.lbl_datos = ttk.Label(txt, text="Elige el PDF que quieres hacer más ligero.", foreground="#555",
                                   wraplength=680, justify="left")
        self.lbl_datos.pack(anchor="w")

        # 2) porcentaje
        opc = ttk.LabelFrame(caja, labelwidget=_etiqueta(caja, "🎚️", "TAMAÑO DESEADO", 24), padding=8)
        opc.pack(fill="x", pady=(10, 0))
        self.btn = _boton(opc, "🗜️", "Comprimir", self.comprimir, px=30)
        self.btn.pack(side="right", padx=(16, 0))
        izq = ttk.Frame(opc)
        izq.pack(side="left", fill="x", expand=True)
        self.lbl_pct = ttk.Label(izq, font=("Segoe UI", 12, "bold"))
        self.lbl_pct.pack(anchor="w")
        fila = ttk.Frame(izq)
        fila.pack(fill="x", pady=(4, 2))
        ttk.Label(fila, text="10%", foreground="#555").pack(side="left")
        self.escala = ttk.Scale(fila, from_=10, to=90, orient="horizontal", command=self._mover)
        self.escala.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Label(fila, text="90%", foreground="#555").pack(side="left")
        ttk.Label(izq, text="Mientras más bajo, más ligero queda el archivo y menor es la calidad de sus imágenes. "
                            "El texto se conserva.", foreground="#555").pack(anchor="w")
        self.pct = int(cfg.get("pct_pdf", 50))
        self.escala.set(self.pct)
        self._mover(self.pct)

        # 3) resultado: antes -> ahora
        res = ttk.LabelFrame(caja, labelwidget=_etiqueta(caja, "📊", "RESULTADO", 24), padding=8)
        res.pack(fill="x", pady=(10, 0))
        self.lbl_antes = self._cifra(res, "Antes", 0)
        ttk.Label(res, image=icono("➡️", 30)).grid(row=0, column=1, rowspan=2, padx=18)
        self.lbl_ahora = self._cifra(res, "Ahora", 2, self.VERDE)
        self.lbl_cambio = ttk.Label(res, text="", font=("Segoe UI", 10), foreground="#555", justify="left")
        self.lbl_cambio.grid(row=0, column=3, rowspan=2, sticky="w", padx=(26, 0))
        res.columnconfigure(4, weight=1)
        self.btn_abrir = _boton(res, "📄", "Abrir PDF", lambda: self.abrir(self.salida), px=22)
        self.btn_abrir.grid(row=0, column=5, rowspan=2, padx=(6, 0))
        self.btn_carpeta = _boton(res, "📁", "Ver en carpeta", lambda: self.carpeta(self.salida), px=22)
        self.btn_carpeta.grid(row=0, column=6, rowspan=2, padx=(6, 0))

        # 4) historial del archivo
        his = ttk.LabelFrame(caja, labelwidget=_etiqueta(caja, "🕘", "HISTORIAL DE ESTE ARCHIVO", 24), padding=8)
        his.pack(fill="both", expand=True, pady=(10, 0))
        acc = ttk.Frame(his)                           # acciones sobre la versión elegida
        acc.pack(side="bottom", fill="x", pady=(8, 0))
        ttk.Label(acc, text="Elige una versión de la lista para abrirla o guardarla donde quieras.",
                  foreground="#555").pack(side="left")
        self.btn_guardar_v = _boton(acc, "💾", "Guardar como…", self.guardar_version, px=20)
        self.btn_guardar_v.pack(side="right")
        self.btn_abrir_v = _boton(acc, "📄", "Abrir", self.abrir_version, px=20)
        self.btn_abrir_v.pack(side="right", padx=6)
        self.btn_quitar_v = _boton(acc, "🗑️", "Quitar del historial", self.quitar_version, px=20)
        self.btn_quitar_v.pack(side="right")
        cols =(("fecha", "Fecha", 120), ("archivo", "Archivo resultante", 290), ("antes", "Antes", 80),
                ("despues", "Después", 80), ("queda", "Quedó en", 75), ("pedido", "Pedido", 65),
                ("calidad", "Calidad", 85))
        self.tabla = ttk.Treeview(his, columns=[c for c, _, _ in cols], show="headings", height=4,
                                  selectmode="browse")
        for c, t, w in cols:
            self.tabla.heading(c, text=t)
            self.tabla.column(c, width=w, minwidth=w, stretch=(c == "archivo"),
                              anchor="w" if c in ("fecha", "archivo") else "center")
        self.tabla.tag_configure("vacio", foreground="#888")
        barra = ttk.Scrollbar(his, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.tabla.bind("<Double-1>", lambda e: self.abrir_version())
        self.tabla.bind("<<TreeviewSelect>>", lambda e: self._botones_historial())

        bot = ttk.Frame(self, padding=(12, 6, 12, 10))
        bot.pack(side="bottom", fill="x", before=caja)   # se reserva primero: nunca queda cortada
        self.barra_estado(bot, 860)

        self._resultado()
        self.refrescar_historial()
        self.msg("ℹ️ Elige el PDF que quieres hacer más ligero.")

    def _cifra(self, padre, titulo, col, color=None):
        ttk.Label(padre, text=titulo, foreground="#555").grid(row=0, column=col, sticky="w")
        l = ttk.Label(padre, text="—", font=("Segoe UI", 20, "bold"), foreground=color)
        l.grid(row=1, column=col, sticky="w")
        return l

    def _mover(self, valor):
        """La barra avanza de 5 en 5 y muestra cuánto pesaría el archivo."""
        v = int(round(float(valor) / 5) * 5)
        self.pct = v
        if abs(float(self.escala.get()) - v) > 0.01:
            self.escala.set(v)
        txt = f"Que pese el {v}% de su tamaño actual"
        if self.tam:
            txt += f"  (≈ {tamano_legible(self.tam * v / 100)})"
        self.lbl_pct.config(text=txt)

    # --- archivo
    def elegir(self):
        if self.trabajando:
            return
        ruta = filedialog.askopenfilename(parent=self, title="Elegir PDF", filetypes=[("PDF", "*.pdf")])
        if ruta:
            self.cargar(ruta)

    def cargar(self, ruta):
        try:
            paginas, imagenes = info_pdf(ruta)
            tam = os.path.getsize(ruta)
        except Exception as e:
            messagebox.showerror("❌ PDF", f"❌ {e}", parent=self)
            return
        self.ruta, self.tam, self.salida = os.path.normpath(ruta), tam, None
        self.lbl_nombre.config(text=os.path.basename(self.ruta))
        carpeta = os.path.dirname(self.ruta)
        if len(carpeta) > 95:                          # ruta muy larga: se recorta en medio
            carpeta = carpeta[:40] + "…" + carpeta[-52:]
        self.lbl_datos.config(text=f"Tamaño actual: {tamano_legible(tam)}  ·  {paginas} página(s)\n"
                                   f"Carpeta: {carpeta}")
        self._mover(self.pct)
        self._resultado()
        self.refrescar_historial()
        if imagenes == 0:
            self.msg("⚠️ Este PDF no tiene imágenes (solo texto o dibujos): se podrá reducir poco, sin perder calidad.")
        else:
            self.msg("ℹ️ Mueve la barra al porcentaje deseado y presiona Comprimir.")

    # --- comprimir
    def comprimir(self):
        if self.trabajando:
            return
        if not self.ruta:
            messagebox.showinfo("ℹ️ Falta algo", "ℹ️ Primero elige un PDF.", parent=self)
            return
        if not os.path.exists(self.ruta):
            messagebox.showerror("❌ PDF", "❌ El archivo ya no existe. Elígelo de nuevo.", parent=self)
            return
        origen, pct = self.ruta, self.pct
        cfg = cargar_config()
        cfg["pct_pdf"] = pct
        guardar_config(cfg)
        self.trabajando = True
        self.btn.state(["disabled"])
        self.msg(f"⏳ Comprimiendo al {pct}%… con archivos grandes puede tardar un poco.")

        def progreso(n):
            self._luego(lambda: self.msg(f"⏳ Comprimiendo al {pct}%… buscando la mejor calidad (prueba {n})"))

        def trabajo():
            try:
                with open(origen, "rb") as f:
                    datos = f.read()
                nuevo, nivel = comprimir_pdf(datos, pct, progreso)
                destino = copia = despues = None
                if nuevo is not None:
                    destino = guardar_comprimido(origen, pct, nuevo)
                    despues = len(nuevo)
                    try:
                        copia = archivar_version(nuevo)
                    except Exception:
                        pass                           # sin copia: el historial usa el archivo junto al original
                self._luego(lambda: self._listo(origen, len(datos), despues, nivel, pct, destino, copia))
            except Exception as e:
                self._luego(lambda m=str(e): self._fallo(m))
        threading.Thread(target=trabajo, daemon=True).start()

    def _listo(self, origen, antes, despues, nivel, pct, destino, copia=None):
        self.trabajando = False
        self.btn.state(["!disabled"])
        if despues is None:
            self.msg("⚠️ Este PDF ya está optimizado: no se pudo hacer más ligero. No se creó ningún archivo.")
            return
        calidad = "sin pérdida" if nivel < 0 else NIVELES_PDF[nivel][2]
        try:
            agregar_historial({"fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "origen": origen,
                               "salida": destino, "copia": copia, "antes": antes, "despues": despues,
                               "pedido": pct, "calidad": calidad})
        except Exception:
            pass                                       # el PDF ya se guardó; el historial no es crítico
        self.salida = destino
        self._resultado(antes, despues, calidad)
        self.refrescar_historial(resaltar=True)
        real = despues / antes * 100
        if real <= pct:
            extra = " Quedó aún más ligero sin bajar más la calidad." if real < pct - 10 and nivel <= 0 else ""
            self.msg(f"✅ Listo: ahora pesa el {real:.0f}% (pediste {pct}%), calidad {calidad}.{extra} "
                     f"Guardado como {os.path.basename(destino)}")
        elif nivel < 0:
            self.msg(f"⚠️ Este PDF casi no tiene imágenes: solo se pudo bajar al {real:.0f}% sin perder calidad. "
                     f"Guardado como {os.path.basename(destino)}")
        else:
            self.msg(f"⚠️ Lo más que se pudo bajar fue al {real:.0f}% (pediste {pct}%). "
                     f"Guardado como {os.path.basename(destino)}")

    def _fallo(self, m):
        self.trabajando = False
        self.btn.state(["!disabled"])
        self.msg("❌ No se pudo comprimir el PDF.")
        messagebox.showerror("❌ Comprimir", "❌ No se pudo comprimir el PDF:\n" + m, parent=self)

    def _resultado(self, antes=None, despues=None, calidad=""):
        """Muestra el antes y el ahora del último PDF comprimido (o solo el tamaño actual)."""
        botones = (self.btn_abrir, self.btn_carpeta)
        if antes is None:
            self.lbl_antes.config(text=tamano_legible(self.tam) if self.tam else "—")
            self.lbl_ahora.config(text="—")
            self.lbl_cambio.config(text="")
            for b in botones:
                b.state(["disabled"])
            return
        self.lbl_antes.config(text=tamano_legible(antes))
        self.lbl_ahora.config(text=tamano_legible(despues))
        self.lbl_cambio.config(text=f"Se redujo {100 - despues / antes * 100:.0f}%\n"
                                    f"Ahora pesa el {despues / antes * 100:.0f}% · calidad {calidad}")
        for b in botones:
            b.state(["!disabled"])

    # --- historial
    def refrescar_historial(self, resaltar=False):
        """Llena la tabla con las versiones de este archivo; resaltar=True selecciona la más reciente."""
        self.tabla.delete(*self.tabla.get_children())
        self.filas = {}
        hist = historial_de(self.ruta) if self.ruta else []
        if not self.ruta:
            self.tabla.insert("", "end", values=("", "Elige un PDF para ver su historial."), tags=("vacio",))
        elif not hist:
            self.tabla.insert("", "end", values=("", "Este archivo todavía no se ha comprimido."), tags=("vacio",))
        for h in sorted(reversed(hist), key=lambda x: x.get("fecha", ""), reverse=True):   # lo más reciente arriba
            try:
                fecha = datetime.strptime(h["fecha"], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
            except Exception:
                fecha = h.get("fecha", "")
            nombre = os.path.basename(h["salida"]) + ("" if version_disponible(h) else "  (ya no existe)")
            iid = self.tabla.insert("", "end", values=(
                fecha, nombre, tamano_legible(h["antes"]), tamano_legible(h["despues"]),
                f"{h['despues'] / h['antes'] * 100:.0f}%" if h.get("antes") else "",
                f"{h['pedido']}%" if h.get("pedido") else "", h.get("calidad", "")))
            self.filas[iid] = h
        if resaltar and self.filas:
            primera = self.tabla.get_children()[0]
            self.tabla.selection_set(primera)
            self.tabla.see(primera)
        self._botones_historial()

    def _version_elegida(self):
        sel = self.tabla.selection()
        return self.filas.get(sel[0]) if sel else None

    def _botones_historial(self):
        h = self._version_elegida()
        hay = h is not None and version_disponible(h) is not None
        self.btn_guardar_v.state(["!disabled"] if hay else ["disabled"])
        self.btn_abrir_v.state(["!disabled"] if hay else ["disabled"])
        self.btn_quitar_v.state(["!disabled"] if h is not None else ["disabled"])

    def guardar_version(self):
        """Guarda la versión elegida del historial donde el usuario quiera."""
        h = self._version_elegida()
        fuente = version_disponible(h) if h else None
        if fuente is None:
            messagebox.showinfo("ℹ️ Historial", "ℹ️ Elige en la lista una versión que todavía exista.", parent=self)
            return
        cfg = cargar_config()
        carpeta = cfg.get("carpeta_guardar_pdf") or os.path.dirname(h["salida"])
        destino = filedialog.asksaveasfilename(
            parent=self, title=f"Guardar la versión de {tamano_legible(h['despues'])}",
            initialdir=carpeta if os.path.isdir(carpeta) else None, initialfile=os.path.basename(h["salida"]),
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not destino:
            return
        if _clave_ruta(destino) != _clave_ruta(fuente):
            try:
                shutil.copyfile(fuente, destino)
            except PermissionError:
                messagebox.showerror("❌ Guardar", "❌ No se pudo guardar: el archivo está abierto en otro programa "
                                                  "o la carpeta no permite escribir.", parent=self)
                return
            except OSError as e:
                messagebox.showerror("❌ Guardar", f"❌ No se pudo guardar:\n{e}", parent=self)
                return
        cfg["carpeta_guardar_pdf"] = os.path.dirname(destino)
        guardar_config(cfg)
        self.msg(f"✅ Versión de {tamano_legible(h['despues'])} guardada como «{os.path.basename(destino)}» "
                 f"en la carpeta «{os.path.basename(os.path.dirname(os.path.normpath(destino)))}».")

    def abrir_version(self):
        h = self._version_elegida()
        if h is not None:
            self.abrir(version_disponible(h))

    def quitar_version(self):
        h = self._version_elegida()
        if h is None:
            return
        if not messagebox.askyesno(
                "Quitar del historial",
                f"¿Quitar la versión de {tamano_legible(h['despues'])} del historial?\n\n"
                f"El archivo «{os.path.basename(h['salida'])}» junto al original no se borra.", parent=self):
            return
        try:
            quitar_de_historial(h)
        except Exception as e:
            messagebox.showerror("❌ Historial", f"❌ No se pudo actualizar el historial:\n{e}", parent=self)
            return
        self.refrescar_historial()
        self.msg("🧹 Versión quitada del historial.")

    def abrir(self, ruta):
        if ruta and os.path.exists(ruta):
            os.startfile(ruta)
        else:
            messagebox.showinfo("ℹ️ PDF", "ℹ️ Ese archivo ya no existe (se movió o se borró).", parent=self)

    def carpeta(self, ruta):
        if ruta and os.path.exists(ruta):
            subprocess.Popen(f'explorer /select,"{os.path.normpath(ruta)}"')
        else:
            messagebox.showinfo("ℹ️ PDF", "ℹ️ Ese archivo ya no existe (se movió o se borró).", parent=self)


class App(tk.Tk):
    """Ventana única. Inicio: título y funciones. Cada función se muestra en la misma ventana, con un
    encabezado (botón Regresar + nombre + instrucciones) y la barra de escáner/papel, compartida por todas."""

    def __init__(self):
        super().__init__()
        self.title("INEScan 2x")
        try:
            self.iconbitmap(default=ICONO_APP)        # icono de la app (también en los diálogos)
        except Exception:
            pass                                       # sin icono no es crítico
        ancho, alto = 920, min(750, self.winfo_screenheight() - 80)
        x = (self.winfo_screenwidth() - ancho) // 2
        y = max(0, (self.winfo_screenheight() - alto) // 2 - 20)
        self.geometry(f"{ancho}x{alto}+{x}+{y}")       # centrada en la pantalla
        self.resizable(False, False)                   # tamaño fijo
        self.escaneres = []
        self.escaneando = False

        st = ttk.Style(self)
        st.configure("Icono.TButton", padding=(10, 6))
        funciones = [("credencial", Credencial), ("oficio", DocOficio), ("varias", VariasPaginas),
                     ("comprimir", ComprimirPDF)]

        # barra superior, siempre visible: el logo lleva al inicio y los enlaces a cada función
        nav = tk.Frame(self, bg="#ffffff", height=50)
        nav.pack(fill="x")
        nav.pack_propagate(False)
        tk.Frame(self, bg="#e4e4ea", height=1).pack(fill="x")
        logo = [tk.Label(nav, image=logo_app(28), bg="#ffffff"),
                tk.Label(nav, text=" INEScan", font=("Segoe UI", -20, "bold"), fg=TINTA, bg="#ffffff"),
                tk.Label(nav, text="2x", font=("Segoe UI", -20, "bold"), fg=ACENTO, bg="#ffffff")]
        for i, p in enumerate(logo):
            p.config(cursor="hand2")
            p.pack(side="left", padx=(16 if i == 0 else 0, 0))
            p.bind("<Button-1>", lambda e: self.ir("inicio"))
        self.enlaces = {}
        for i, (clave, cls) in enumerate(funciones):
            l = tk.Label(nav, text=cls.TITULO.upper(), font=("Segoe UI", -11, "bold"), fg=TINTA,
                         bg="#ffffff", cursor="hand2")
            l.pack(side="left", padx=(34 if i == 0 else 22, 0))
            l.bind("<Button-1>", lambda e, c=clave: self.ir(c))
            l.bind("<Enter>", lambda e, l=l: l.config(fg=ACENTO))
            l.bind("<Leave>", lambda e: self._pintar_enlaces())
            self.enlaces[clave] = l

        # encabezado de las funciones (oculto en el inicio)
        self.cab = ttk.Frame(self, padding=(12, 10, 12, 0))
        _boton(self.cab, "⬅️", "Regresar", lambda: self.ir("inicio")).pack(side="left")
        self.cab_titulo = ttk.Label(self.cab, compound="left", font=("Segoe UI", 15, "bold"))
        self.cab_titulo.pack(side="left", padx=(16, 0))
        self.cab_ayuda = ttk.Label(self.cab, wraplength=470, foreground="#555", justify="left")
        self.cab_ayuda.pack(side="left", padx=16)

        # escáner y papel, compartidos por todas las funciones (oculto en el inicio)
        self.barra = ttk.Frame(self, padding=(12, 10, 12, 0))
        _etiqueta(self.barra, "🖨️", "Escáner:").pack(side="left")
        self.cmb = ttk.Combobox(self.barra, state="readonly", width=40)
        self.cmb.pack(side="left", padx=6)
        _boton(self.barra, "🔄", "Buscar", self.buscar).pack(side="left")
        _boton(self.barra, "🌐", "Red…", self.agregar_red).pack(side="left", padx=(6, 0))
        self.cmb.bind("<<ComboboxSelected>>", self.recordar_escaner)
        _etiqueta(self.barra, "📄", "Papel:").pack(side="left", padx=(14, 4))
        self.papel = ttk.Combobox(self.barra, state="readonly", width=7, values=["Carta", "A4"])
        self.papel.current(0)
        self.papel.pack(side="left")

        self.cuerpo = ttk.Frame(self)
        self.cuerpo.pack(fill="both", expand=True)
        self.vistas = {clave: cls(self) for clave, cls in funciones}
        self.vistas["inicio"] = Inicio(self, funciones)
        self.actual = None
        self.ir("inicio")
        self.after(200, self.buscar)

    def ir(self, nombre):
        """Muestra una pantalla y oculta la anterior (cada función conserva lo que ya se escaneó)."""
        if self.actual is not None:
            self.actual.pack_forget()
        v = self.vistas[nombre]
        if nombre == "inicio":
            self.cab.pack_forget()
            self.barra.pack_forget()
            self.title("INEScan 2x")
        else:
            self.cab_titulo.config(text=" " + v.TITULO, image=icono_plano(v.GLIFO, v.COLOR, 34, v.EMOJI))
            self.cab_ayuda.config(text=v.AYUDA)
            self.cab.pack(fill="x", before=self.cuerpo)
            if v.USA_ESCANER:
                self.barra.pack(fill="x", before=self.cuerpo)
            else:
                self.barra.pack_forget()
            self.title(f"INEScan 2x · {v.TITULO}")
        v.pack(fill="both", expand=True)
        self.actual = v
        v.al_mostrar()
        self._pintar_enlaces()

    def _pintar_enlaces(self):
        """El enlace de la función visible va en color de acento."""
        for clave, l in self.enlaces.items():
            l.config(fg=ACENTO if self.actual is self.vistas.get(clave) else TINTA)

    def msg(self, t):
        """Mensaje en la barra de estado de la pantalla visible."""
        if self.actual is not None:
            self.actual.msg(t)

    def pagina(self):
        return letter if self.papel.get() == "Carta" else A4

    # --- escáner
    def buscar(self):
        err = ""
        try:
            self.escaneres = listar_escaneres()
        except Exception as e:
            self.escaneres = []
            err = str(e)
        for h in cargar_config().get("redes", []):            # impresoras de red agregadas antes
            self.escaneres.append((f"Red: {h}", "escl:" + h))
        self.cmb["values"] = [n for n, _ in self.escaneres]
        if self.escaneres:
            ultimo = cargar_config().get("ultimo")
            nombres = [n for n, _ in self.escaneres]
            if ultimo in nombres:                              # el que se usó la última vez
                idx = nombres.index(ultimo)
            else:                                              # si no, uno con marca conocida o el primero
                idx = next((i for i, n in enumerate(nombres) if "HP" in n.upper() or "SMART" in n.upper()), 0)
            self.cmb.current(idx)
            self.msg(f"✅ Escáner listo: {nombres[idx]}")
        elif err:
            self.msg(f"❌ Error buscando escáner: {err}")
        else:
            self.msg("⚠️ No se encontró escáner. Enciende la impresora/conéctala, usa 'Red…' o 'Abrir imagen…'.")

    def recordar_escaner(self, _=None):
        i = self.cmb.current()
        if i >= 0:
            cfg = cargar_config()
            cfg["ultimo"] = self.escaneres[i][0]
            guardar_config(cfg)

    def agregar_red(self):
        """Agrega una impresora/escáner de red por su IP (eSCL / AirScan), para equipos que Windows no lista."""
        host = simpledialog.askstring(
            "Escáner de red",
            "Dirección IP o nombre de la impresora (ej. 192.168.1.50).\n"
            "Debe estar encendida y en la misma red. La IP aparece en el panel de la impresora.",
            parent=self)
        if not host or not host.strip():
            return
        host = host.strip()
        self.msg("⏳ Buscando el escáner de red…")

        def trabajo():
            try:
                _escl_base(host)
                self.after(0, lambda: self._red_lista(host))
            except Exception as e:
                self.after(0, lambda m=str(e): (self.msg("❌ No se encontró el escáner de red."),
                                                messagebox.showerror("❌ Red", m, parent=self)))
        threading.Thread(target=trabajo, daemon=True).start()

    def _red_lista(self, host):
        cfg = cargar_config()
        redes = cfg.get("redes", [])
        if host not in redes:
            redes.append(host)
        cfg["redes"] = redes
        cfg["ultimo"] = f"Red: {host}"
        guardar_config(cfg)
        self.buscar()

    def escanear_async(self, al_terminar, msg, espera):
        """Escanea con el escáner elegido sin congelar la ventana y llama al_terminar(imagen).
        Solo un escaneo a la vez, aunque se cambie de pantalla mientras tanto."""
        if self.escaneando:
            messagebox.showinfo("⏳ Escáner", "⏳ El escáner está ocupado. Espera a que termine el escaneo.",
                                parent=self)
            return
        i = self.cmb.current()
        if i < 0:
            messagebox.showwarning("⚠️ Escáner", "⚠️ No hay escáner seleccionado.\n"
                                                 "Presiona 🔄 Buscar o usa 📂 Abrir imagen.", parent=self)
            return
        dev_id = self.escaneres[i][1]
        self.escaneando = True
        msg(espera)

        def listo(im):
            self.escaneando = False
            al_terminar(im)

        def fallo(m):
            self.escaneando = False
            msg("❌ Falló el escaneo.")
            messagebox.showerror("❌ Escaneo", f"❌ No se pudo escanear:\n{m}", parent=self)

        def trabajo():
            try:
                import pythoncom
                pythoncom.CoInitialize()
                im = escanear(dev_id)
                self.after(0, lambda: listo(im))
            except Exception as e:
                self.after(0, lambda m=str(e): fallo(m))
        threading.Thread(target=trabajo, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
