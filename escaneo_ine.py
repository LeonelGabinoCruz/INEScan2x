"""
INEScan 2x - Escaneo de INE / credencial (frente y reverso) -> PDF al DOBLE de tamaño, listo para imprimir.
Pensado para HP Smart Tank 525 (usa WIA de Windows).

Flujo:
  1. Pon la credencial por el FRENTE en el cristal -> "Escanear frente".
  2. Voltéala -> "Escanear reverso".
  3. "PDF al doble (2x)" o "PDF tamaño normal" -> se recorta solo y se abre el PDF.
"""
import json
import os
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


# ----------------------------------------------------------------- interfaz
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("INEScan 2x")
        try:
            self.iconbitmap(default=ICONO_APP)        # icono de la app (también en las ventanas hijas)
        except Exception:
            pass                                       # sin icono no es crítico
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
        boton(cab, "📑", "Documento oficio…", lambda: DocOficio(self)).pack(side="right")

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        etiqueta(top, "🖨️", "Escáner:").pack(side="left")
        self.cmb = ttk.Combobox(top, state="readonly", width=40)
        self.cmb.pack(side="left", padx=6)
        boton(top, "🔄", "Buscar", self.buscar).pack(side="left")
        boton(top, "🌐", "Red…", self.agregar_red).pack(side="left", padx=(6, 0))
        self.cmb.bind("<<ComboboxSelected>>", self.recordar_escaner)
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
            self.msg("✅ Escáner listo.")
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
                                                messagebox.showerror("❌ Red", m)))
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
                im = escanear(dev_id)
                self.after(0, lambda: self.procesar(clave, im))
            except Exception as e:
                self.after(0, lambda m=str(e): (self.msg("❌ Falló el escaneo."),
                                                messagebox.showerror("❌ Escaneo", f"❌ No se pudo escanear:\n{m}")))
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


def _boton(padre, emoji, texto, cmd, px=26):
    return ttk.Button(padre, text=" " + texto, image=icono(emoji, px),
                      compound="left", command=cmd, style="Icono.TButton")


def _etiqueta(padre, emoji, texto, px=24, **kw):
    return ttk.Label(padre, text=" " + texto, image=icono(emoji, px), compound="left", **kw)


class DocOficio(tk.Toplevel):
    """Documento más largo que el cristal: se escanea en 2 partes, se unen solas y se elige cómo imprimir."""
    FORMAS = [("oficio", "Oficio, tamaño real (21.6 x 34 cm)"),
              ("carta", "Ajustar a 1 hoja carta/A4 (reducido)"),
              ("dividir", "Dividir en hojas carta/A4 (tamaño real)")]

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Documento oficio en 2 partes")
        ancho, alto = 920, 640
        self.geometry(f"{ancho}x{alto}+{(self.winfo_screenwidth() - ancho) // 2}+{(self.winfo_screenheight() - alto) // 2}")
        self.resizable(False, False)
        self.transient(app)
        self.partes = {"1": None, "2": None}
        self.unido = None

        cab = ttk.Frame(self, padding=(12, 10, 12, 0))
        cab.pack(fill="x")
        ttk.Label(cab, text=" Documento oficio", image=icono("📑", 36), compound="left",
                  font=("Segoe UI", 15, "bold")).pack(side="left")
        ttk.Label(cab, text="1) Escanea la parte de ARRIBA.  2) Desliza el documento y escanea lo que falta, "
                            "dejando ~2 cm repetidos.  3) Elige cómo imprimir.",
                  wraplength=600, foreground="#555", justify="left").pack(side="left", padx=14)

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
        self.ico_estado = ttk.Label(est)
        self.ico_estado.pack(side="left")
        self.estado = ttk.Label(est, text="", wraplength=780, justify="left")
        self.estado.pack(side="left", padx=6)
        self.msg("ℹ️ Escanea la parte 1 (la de ARRIBA del documento).")

    def msg(self, t):
        """El primer símbolo del mensaje (✅ ⚠️ ❌ ⏳ ℹ️ ...) se muestra como icono grande."""
        t = t.strip()
        e_ini = ""
        for e in ("✅", "⚠️", "❌", "⏳", "ℹ️", "🧹"):
            if t.startswith(e):
                e_ini, t = e, t[len(e):].strip()
                break
        self.ico_estado.config(image=icono(e_ini, 30) if e_ini else "")
        self.estado.config(text=t)
        self.update_idletasks()

    def mostrar(self, clave, im):
        mini = im.copy()
        mini.thumbnail((260, 320))
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
            messagebox.showerror("Unir", f"No se pudo unir:\n{e}")
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
        i = self.app.cmb.current()
        if i < 0:
            messagebox.showwarning("Escáner", "No hay escáner seleccionado en la ventana principal.")
            return
        dev_id = self.app.escaneres[i][1]
        self.msg(f"⏳ Escaneando parte {clave}… espera.")

        def trabajo():
            try:
                import pythoncom
                pythoncom.CoInitialize()
                im = escanear(dev_id)
                self.after(0, lambda: self.procesar(clave, im))
            except Exception as e:
                self.after(0, lambda m=str(e): (self.msg("❌ Falló el escaneo."),
                                                messagebox.showerror("Escaneo", f"No se pudo escanear:\n{m}")))
        threading.Thread(target=trabajo, daemon=True).start()

    def abrir(self, clave):
        ruta = filedialog.askopenfilename(filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff")])
        if ruta:
            src = Image.open(ruta)
            d = src.info.get("dpi", (DPI, DPI))[0] or DPI
            im = src.convert("RGB")
            im.dpi_real = d
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
            messagebox.showinfo("Falta algo", "Escanea primero la parte 1 y la parte 2.")
            return
        modo = self.FORMAS[self.forma.current()][0]
        os.makedirs(CARPETA, exist_ok=True)
        ruta = os.path.join(CARPETA, datetime.now().strftime("documento_%Y%m%d_%H%M%S") + f"_{modo}.pdf")
        pagina = letter if self.app.papel.get() == "Carta" else A4
        try:
            generar_pdf_documento(self.unido, ruta, modo, pagina)
        except Exception as e:
            messagebox.showerror("PDF", "No se pudo crear el PDF:\n" + str(e))
            return
        self.msg(f"✅ PDF guardado: {ruta}")
        os.startfile(ruta)


if __name__ == "__main__":
    App().mainloop()
