"""
Evaluador ROSA — NTP 1173 (INSST, 2022)   v4.1
================================================
Compatibilidad: MediaPipe 0.10.x (nueva API)
"""

import cv2
import mediapipe as mp
from mediapipe.tasks import python as mp_python
from mediapipe.tasks.python import vision as mp_vision
from mediapipe.tasks.python.components import containers
import numpy as np
import time, os, threading, queue, urllib.request, sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk

from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side)
from openpyxl.utils import get_column_letter

try:
    import winsound
    WINSOUND_OK = True
except ImportError:
    WINSOUND_OK = False

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────
INTERVALO_EVAL_SEG = 10
GUARDADO_SEG       = 3
MAX_CAMARAS_SCAN   = 6
NIVEL_ACCION       = 5

# URL del modelo de pose landmarker de MediaPipe 0.10+
MODEL_URL  = "https://storage.googleapis.com/mediapipe-models/pose_landmarker/pose_landmarker_lite/float16/latest/pose_landmarker_lite.task"
MODEL_PATH = os.path.join(os.path.expanduser("~"), "pose_landmarker_lite.task")

C = {
    "bg_deep":   "#0A0C10",
    "bg_panel":  "#111318",
    "bg_card":   "#181C22",
    "bg_border": "#252A34",
    "accent":    "#00C8FF",
    "accent2":   "#0087CC",
    "text_hi":   "#F0F4FF",
    "text_mid":  "#9BAEC8",
    "text_lo":   "#56657A",
    "ok":        "#00D68F",
    "warn":      "#F6AD3B",
    "danger":    "#FF4757",
}

NIVELES = {
    1: ("Inapreciable",        "#00D68F"),
    2: ("Bajo",                "#00B87A"),
    3: ("Medio",               "#F6AD3B"),
    4: ("Medio-Alto",          "#E8823A"),
    5: ("ALTO — ACTUAR YA",   "#FF4757"),
}

# Índices de landmarks en MediaPipe 0.10+ PoseLandmarker
# (mismos números que antes pero acceso diferente)
_LM = {
    "NOSE":            0,
    "LEFT_SHOULDER":  11,
    "RIGHT_SHOULDER": 12,
    "LEFT_ELBOW":     13,
    "LEFT_WRIST":     15,
    "LEFT_HIP":       23,
    "RIGHT_HIP":      24,
    "LEFT_KNEE":      25,
    "LEFT_ANKLE":     27,
}

# ─────────────────────────────────────────────────────────────
# TABLAS NTP 1173
# ─────────────────────────────────────────────────────────────
_TABLA_A = [
    [2,2,3,4,5,6,7,8],
    [2,2,3,4,5,6,7,8],
    [3,3,3,4,5,6,7,8],
    [4,4,4,4,5,6,7,8],
    [5,5,5,5,6,7,8,9],
    [6,6,6,7,7,8,8,9],
    [7,7,7,8,8,9,9,9],
]
_TABLA_B = [
    [1,1,1,2,3,4,5,6,6],
    [1,1,2,2,3,4,5,6,6],
    [1,2,2,3,3,4,6,7,7],
    [2,2,3,3,4,5,6,8,8],
    [3,3,4,4,5,6,7,8,8],
    [4,4,5,5,6,7,8,9,9],
    [5,5,6,7,8,8,9,9,9],
]
_TABLA_C = [
    [1,1,1,2,3,4,5,6],
    [1,1,2,3,4,5,6,7],
    [1,2,2,3,4,5,6,7],
    [2,3,3,3,5,6,7,8],
    [3,4,4,5,5,6,7,8],
    [4,5,5,6,6,7,8,9],
    [5,6,6,7,7,8,8,9],
    [6,7,7,8,8,9,9,9],
]
_TABLA_D = [
    [1,2,3,4,5,6,7,8,9],
    [2,2,3,4,5,6,7,8,9],
    [3,3,3,4,5,6,7,8,9],
    [4,4,4,4,5,6,7,8,9],
    [5,5,5,5,5,6,7,8,9],
    [6,6,6,6,6,6,7,8,9],
    [7,7,7,7,7,7,7,8,9],
    [8,8,8,8,8,8,8,8,9],
    [9,9,9,9,9,9,9,9,9],
]
_TABLA_E = [
    [1, 2, 3, 4, 5, 6, 7, 8, 9,10],
    [2, 2, 3, 4, 5, 6, 7, 8, 9,10],
    [3, 3, 3, 4, 5, 6, 7, 8, 9,10],
    [4, 4, 4, 4, 5, 6, 7, 8, 9,10],
    [5, 5, 5, 5, 5, 6, 7, 8, 9,10],
    [6, 6, 6, 6, 6, 6, 7, 8, 9,10],
    [7, 7, 7, 7, 7, 7, 7, 8, 9,10],
    [8, 8, 8, 8, 8, 8, 8, 8, 9,10],
    [9, 9, 9, 9, 9, 9, 9, 9, 9,10],
    [10,10,10,10,10,10,10,10,10,10],
]

def _tlu(tabla, fi, ci):
    fi = int(np.clip(fi, 0, len(tabla) - 1))
    ci = int(np.clip(ci, 0, len(tabla[0]) - 1))
    return int(tabla[fi][ci])

# ─────────────────────────────────────────────────────────────
# LÓGICA ROSA
# ─────────────────────────────────────────────────────────────
def factor_tiempo_F(horas_diarias):
    if horas_diarias > 4:   return +1
    elif horas_diarias < 1: return -1
    else:                   return 0

def puntuar_A1(ang_rodilla):
    return 1 if 85 <= ang_rodilla <= 100 else 2

def puntuar_A2(dist=8):
    return 1 if 5 <= dist <= 9 else 2

def puntuar_A3(ang_codo):
    return 1 if 80 <= ang_codo <= 110 else 2

def puntuar_A4(ang_tronco):
    incl = 90 + ang_tronco
    return 1 if 95 <= incl <= 110 else 2

def puntuar_B1():
    return 1  # asume manos libres y distancia correcta

def puntuar_B2(desv_cuello, factor_t=0):
    if abs(desv_cuello) < 10:   p = 1
    elif desv_cuello < 0:       p = 2
    else:                       p = 3
    return int(np.clip(p + factor_t, 1, 9))

def puntuar_C1(factor_t=0):
    return int(np.clip(1 + factor_t, 1, 9))

def puntuar_C2(ang_muneca, factor_t=0):
    p = 1 if abs(ang_muneca) <= 15 else 2
    return int(np.clip(p + factor_t, 1, 9))

def calcular_ROSA_completo(ang_tronco, ang_rodilla, ang_codo,
                            desv_cuello, ang_muneca, horas_diarias):
    ft = factor_tiempo_F(horas_diarias)
    a1, a2 = puntuar_A1(ang_rodilla), puntuar_A2()
    a3, a4 = puntuar_A3(ang_codo), puntuar_A4(ang_tronco)
    suma_asiento = a1 + a2
    suma_soporte = a3 + a4
    tabla_A_val  = _tlu(_TABLA_A, suma_asiento - 2, suma_soporte - 2)
    total_silla  = int(np.clip(tabla_A_val + ft, 1, 10))

    b1 = puntuar_B1()
    b2 = puntuar_B2(desv_cuello, ft)
    tabla_B_val = _tlu(_TABLA_B, b1 - 1, b2 - 1)

    c1 = puntuar_C1(ft)
    c2 = puntuar_C2(ang_muneca, ft)
    tabla_C_val = _tlu(_TABLA_C, c1 - 1, c2 - 1)

    tabla_D_val = _tlu(_TABLA_D, tabla_B_val - 1, tabla_C_val - 1)
    rosa = _tlu(_TABLA_E, total_silla - 1, tabla_D_val - 1)

    return {
        "ang_tronco":  round(ang_tronco, 1),
        "ang_rodilla": round(ang_rodilla, 1),
        "ang_codo":    round(ang_codo, 1),
        "desv_cuello": round(desv_cuello, 1),
        "ang_muneca":  round(ang_muneca, 1),
        "A1": a1, "A2": a2, "A3": a3, "A4": a4,
        "suma_asiento": suma_asiento, "suma_soporte": suma_soporte,
        "tabla_A": tabla_A_val, "factor_tiempo": ft,
        "total_silla": total_silla,
        "B1": b1, "B2": b2, "tabla_B": tabla_B_val,
        "C1": c1, "C2": c2, "tabla_C": tabla_C_val,
        "tabla_D": tabla_D_val, "rosa": rosa,
    }

# ─────────────────────────────────────────────────────────────
# ÁNGULOS
# ─────────────────────────────────────────────────────────────
def calcular_angulo(a, b, c):
    a, b, c = np.array(a), np.array(b), np.array(c)
    ba, bc = a - b, c - b
    cos = np.dot(ba, bc) / (np.linalg.norm(ba) * np.linalg.norm(bc) + 1e-6)
    return float(np.degrees(np.arccos(np.clip(cos, -1, 1))))

def extraer_angulos_v2(landmarks, w, h):
    """
    Extrae ángulos usando la nueva API de MediaPipe 0.10+
    landmarks: lista de NormalizedLandmark
    """
    try:
        def gp(idx):
            lm = landmarks[idx]
            return [lm.x * w, lm.y * h]

        lsh  = gp(_LM["LEFT_SHOULDER"])
        rsh  = gp(_LM["RIGHT_SHOULDER"])
        lhi  = gp(_LM["LEFT_HIP"])
        lkn  = gp(_LM["LEFT_KNEE"])
        lank = gp(_LM["LEFT_ANKLE"])
        lelb = gp(_LM["LEFT_ELBOW"])
        lwri = gp(_LM["LEFT_WRIST"])
        nose = gp(_LM["NOSE"])

        ang_tronco  = calcular_angulo(lsh, lhi, [lhi[0], lhi[1] + 100]) - 90
        ang_rodilla = calcular_angulo(lhi, lkn, lank)
        ang_codo    = calcular_angulo(lsh, lelb, lwri)

        mid_sh = [(lsh[0] + rsh[0]) / 2, (lsh[1] + rsh[1]) / 2]
        desv_cuello = calcular_angulo(nose, mid_sh,
                                      [mid_sh[0], mid_sh[1] + 100]) - 90

        ang_muneca = calcular_angulo(lelb, lwri,
                                     [lwri[0] + 100, lwri[1]]) - 90

        return ang_tronco, ang_rodilla, ang_codo, desv_cuello, ang_muneca
    except Exception as e:
        print(f"[extraer_angulos] error: {e}")
        return None

# ─────────────────────────────────────────────────────────────
# DESCARGA DEL MODELO
# ─────────────────────────────────────────────────────────────
def descargar_modelo(callback_status=None):
    if os.path.exists(MODEL_PATH):
        print(f"[modelo] Ya existe: {MODEL_PATH}")
        return True
    try:
        if callback_status:
            callback_status("Descargando modelo de pose (~6 MB)...")
        print(f"[modelo] Descargando desde {MODEL_URL}")
        urllib.request.urlretrieve(MODEL_URL, MODEL_PATH)
        print(f"[modelo] Guardado en {MODEL_PATH}")
        return True
    except Exception as e:
        print(f"[modelo] Error al descargar: {e}")
        return False

# ─────────────────────────────────────────────────────────────
# EXCEL
# ─────────────────────────────────────────────────────────────
HEADER_COLS = [
    ("Timestamp",14),("ROSA Final",11),("Nivel",18),("Total Silla",12),
    ("Total Periféric.",13),("A1-Altura",11),("A2-Profund.",12),
    ("A3-Reposabr.",13),("A4-Respaldo",12),("Tabla A",10),
    ("FactorTiempo",13),("B1-Teléfono",12),("B2-Pantalla",12),
    ("Tabla B",10),("C1-Ratón",11),("C2-Teclado",12),("Tabla C",10),
    ("Tabla D",10),("Áng.Tronco°",13),("Áng.Rodilla°",13),
    ("Áng.Codo°",12),("Desv.Cuello°",13),("Áng.Muñeca°",13),
]
COLOR_NIVEL = {1:"00D68F",2:"00B87A",3:"F6AD3B",4:"E8823A",5:"FF4757"}

def _border():
    t = Side(style="thin", color="C0C0C0")
    return Border(left=t, right=t, top=t, bottom=t)

def _hdr_fill():
    return PatternFill("solid", fgColor="1F3864")

def nivel_texto(score):
    return NIVELES[min(max(score,1),5)][0]

def crear_excel_nuevo(ruta):
    wb = Workbook()
    ws = wb.active
    ws.title = "Registros ROSA"
    ws.merge_cells("A1:W1")
    c = ws["A1"]
    c.value = "EVALUACIÓN ERGONÓMICA — MÉTODO ROSA (NTP 1173 INSST 2022)"
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="0A2342")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:W2")
    s = ws["A2"]
    s.value = f"Generado: {time.strftime('%Y-%m-%d %H:%M:%S')}  |  Nivel de acción: >= 5"
    s.font = Font(name="Arial", size=9, italic=True, color="7F9FBF")
    s.fill = PatternFill("solid", fgColor="0D1B2A")
    s.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 16

    for ci, (titulo, ancho) in enumerate(HEADER_COLS, 1):
        cell = ws.cell(row=3, column=ci, value=titulo)
        cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        cell.fill = _hdr_fill()
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = _border()
        ws.column_dimensions[get_column_letter(ci)].width = ancho
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "A4"

    ws2 = wb.create_sheet("Resumen")
    ws2["A1"] = "RESUMEN EVALUACIÓN ROSA"
    ws2["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws2["A1"].fill = PatternFill("solid", fgColor="0A2342")
    ws2["A1"].alignment = Alignment(horizontal="center")
    ws2.merge_cells("A1:D1")
    for ci, h in enumerate(["Indicador","Valor","Descripción"], 1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = _hdr_fill()
        c.alignment = Alignment(horizontal="center")
    ws2.column_dimensions["A"].width = 25
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 35
    wb.save(ruta)
    return wb

def agregar_registro_excel(ruta, registro, num_fila):
    from openpyxl import load_workbook as _lw
    wb = _lw(ruta)
    ws = wb["Registros ROSA"]
    rosa = registro["rosa"]
    nivel = nivel_texto(rosa)
    color_nivel = COLOR_NIVEL.get(min(rosa, 5), "FF4757")
    fila = num_fila + 3
    valores = [
        registro.get("time",""), rosa, nivel,
        registro.get("total_silla",""), registro.get("tabla_D",""),
        registro.get("A1",""), registro.get("A2",""),
        registro.get("A3",""), registro.get("A4",""),
        registro.get("tabla_A",""), registro.get("factor_tiempo",""),
        registro.get("B1",""), registro.get("B2",""),
        registro.get("tabla_B",""), registro.get("C1",""),
        registro.get("C2",""), registro.get("tabla_C",""),
        registro.get("tabla_D",""), registro.get("ang_tronco",""),
        registro.get("ang_rodilla",""), registro.get("ang_codo",""),
        registro.get("desv_cuello",""), registro.get("ang_muneca",""),
    ]
    fondo = "F0F6FF" if num_fila % 2 == 0 else "FFFFFF"
    for ci, val in enumerate(valores, 1):
        cell = ws.cell(row=fila, column=ci, value=val)
        cell.font = Font(name="Arial", size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _border()
        if ci == 2:
            cell.fill = PatternFill("solid", fgColor=color_nivel)
            cell.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        elif ci == 3:
            cell.fill = PatternFill("solid", fgColor=color_nivel)
            cell.font = Font(name="Arial", bold=True, size=9,
                             color="FFFFFF" if rosa >= 3 else "000000")
        else:
            cell.fill = PatternFill("solid", fgColor=fondo)
    ws.row_dimensions[fila].height = 18

    ws2 = wb["Resumen"]
    rosas = []
    for r in range(4, fila + 1):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, (int, float)):
            rosas.append(v)
    resumen_data = [
        ("Total registros", num_fila, "Evaluaciones realizadas"),
        ("Último ROSA", rosa, nivel),
        ("ROSA promedio", round(np.mean(rosas),2) if rosas else "", "Media"),
        ("ROSA máximo", max(rosas) if rosas else "", "Puntuación más alta"),
        ("ROSA mínimo", min(rosas) if rosas else "", "Puntuación más baja"),
        ("Alertas (>=5)", sum(1 for r in rosas if r >= 5),
         "Evaluaciones sobre nivel de acción"),
    ]
    for ri, (ind, val, desc) in enumerate(resumen_data, 3):
        ws2.cell(row=ri, column=1, value=ind).font = Font(name="Arial", bold=True, size=10)
        ws2.cell(row=ri, column=2, value=val).font = Font(name="Arial", size=10)
        ws2.cell(row=ri, column=3, value=desc).font = Font(name="Arial", size=9, color="606060")
    wb.save(ruta)

# ─────────────────────────────────────────────────────────────
# HILO CÁMARA — API MediaPipe 0.10+
# ─────────────────────────────────────────────────────────────
class HiloCamara(threading.Thread):
    def __init__(self, cola_datos, cola_frames, cola_angulos, horas_diarias, cam_idx):
        super().__init__(daemon=True)
        self.cola_datos    = cola_datos
        self.cola_frames   = cola_frames
        self.cola_angulos  = cola_angulos   # ángulos en tiempo real → GUI
        self.horas_diarias = horas_diarias
        self.cam_idx       = cam_idx
        self.activo        = True
        self._buf = {"tronco":[], "rodilla":[], "codo":[],
                     "cuello":[], "muneca":[]}
        # Resultado compartido entre callback y loop
        self._ultimo_resultado = None
        self._lock = threading.Lock()

    def run(self):
        print("[hilo] Iniciando hilo de cámara...")

        # ── Verificar modelo ──
        if not os.path.exists(MODEL_PATH):
            self.cola_datos.put({"error":
                f"Modelo no encontrado en:\n{MODEL_PATH}\n"
                "Ejecuta el programa con internet para descargarlo."})
            return

        # ── Crear PoseLandmarker con callback (modo LIVE_STREAM) ──
        def resultado_callback(result, output_image, timestamp_ms):
            if result.pose_landmarks and len(result.pose_landmarks) > 0:
                with self._lock:
                    self._ultimo_resultado = result.pose_landmarks[0]

        base_opts = mp_python.BaseOptions(model_asset_path=MODEL_PATH)
        opts = mp_vision.PoseLandmarkerOptions(
            base_options=base_opts,
            running_mode=mp_vision.RunningMode.LIVE_STREAM,
            num_poses=1,
            min_pose_detection_confidence=0.5,
            min_pose_presence_confidence=0.5,
            min_tracking_confidence=0.5,
            result_callback=resultado_callback,
        )

        print("[hilo] Cargando modelo de pose...")
        try:
            landmarker = mp_vision.PoseLandmarker.create_from_options(opts)
        except Exception as e:
            self.cola_datos.put({"error": f"Error cargando modelo:\n{e}"})
            return

        print("[hilo] Modelo cargado. Abriendo cámara...")

        # ── Abrir cámara ──
        cap = cv2.VideoCapture(self.cam_idx, cv2.CAP_DSHOW)
        if not cap.isOpened():
            cap = cv2.VideoCapture(self.cam_idx)
        if not cap.isOpened():
            self.cola_datos.put({"error": "No se pudo acceder a la cámara."})
            return

        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
        print("[hilo] Cámara abierta correctamente.")

        last_eval = time.time()
        ts_ms = 0  # timestamp incremental para LIVE_STREAM

        # Conexiones para dibujar esqueleto manualmente
        CONEXIONES = [
            (11,12),(11,13),(13,15),(12,14),(14,16),  # brazos
            (11,23),(12,24),(23,24),                   # torso
            (23,25),(25,27),(24,26),(26,28),           # piernas
            (0,11),(0,12),                             # cuello aprox
        ]

        while self.activo:
            ret, frame = cap.read()
            if not ret or frame is None:
                continue

            frame = cv2.flip(frame, 1)
            h, w, _ = frame.shape

            # Convertir a MediaPipe Image
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            mp_image = mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb)

            # Detectar (asíncrono con callback)
            ts_ms += 33  # ~30fps
            try:
                landmarker.detect_async(mp_image, ts_ms)
            except Exception as e:
                print(f"[hilo] detect_async error: {e}")

            # Dibujar landmarks del último resultado disponible
            with self._lock:
                lms = self._ultimo_resultado

            if lms is not None:
                # Dibujar puntos
                for lm in lms:
                    cx, cy = int(lm.x * w), int(lm.y * h)
                    cv2.circle(frame, (cx, cy), 4, (0, 200, 255), -1)
                # Dibujar huesos
                for i, j in CONEXIONES:
                    if i < len(lms) and j < len(lms):
                        p1 = (int(lms[i].x*w), int(lms[i].y*h))
                        p2 = (int(lms[j].x*w), int(lms[j].y*h))
                        cv2.line(frame, p1, p2, (255, 255, 0), 2)

                # Acumular ángulos + enviar en tiempo real a GUI
                ang = extraer_angulos_v2(lms, w, h)
                if ang:
                    at, ar, ac, dc, am = ang
                    self._buf["tronco"].append(at)
                    self._buf["rodilla"].append(ar)
                    self._buf["codo"].append(ac)
                    self._buf["cuello"].append(dc)
                    self._buf["muneca"].append(am)

                    # Enviar ángulos actuales a la GUI (sin bloquear)
                    if self.cola_angulos.full():
                        try: self.cola_angulos.get_nowait()
                        except queue.Empty: pass
                    self.cola_angulos.put({
                        "at": at, "ar": ar, "ac": ac,
                        "dc": dc, "am": am
                    })

                    # Dibujar ángulos sobre el frame de video
                    def color_ang(val, ok_min, ok_max):
                        return (0,214,143) if ok_min <= abs(val) <= ok_max else (60,80,255)

                    overlay_data = [
                        (f"Tronco:  {at:+.1f}°",  color_ang(at, 0, 15),   20),
                        (f"Rodilla: {ar:.1f}°",    color_ang(ar, 85, 100), 38),
                        (f"Codo:    {ac:.1f}°",    color_ang(ac, 80, 110), 56),
                        (f"Cuello:  {dc:+.1f}°",   color_ang(dc, 0, 10),   74),
                        (f"Muneca:  {am:+.1f}°",   color_ang(am, 0, 15),   92),
                    ]
                    # Fondo semitransparente para el bloque de texto
                    cv2.rectangle(frame, (w-200, 8), (w-4, 102),
                                  (10, 12, 16), -1)
                    cv2.rectangle(frame, (w-200, 8), (w-4, 102),
                                  (37, 42, 52), 1)
                    for texto, color, y in overlay_data:
                        cv2.putText(frame, texto, (w-196, y),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.42,
                                    color, 1, cv2.LINE_AA)

            # Overlay título
            cv2.rectangle(frame, (0,0), (320,20), (10,12,16), -1)
            cv2.putText(frame, "ROSA NTP-1173 Monitor",
                        (8,14), cv2.FONT_HERSHEY_SIMPLEX, 0.5,
                        (0,200,255), 1, cv2.LINE_AA)

            # Enviar frame a GUI
            if self.cola_frames.full():
                try: self.cola_frames.get_nowait()
                except queue.Empty: pass
            self.cola_frames.put(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))

            # Evaluación periódica
            if time.time() - last_eval >= INTERVALO_EVAL_SEG:
                buf = self._buf
                if all(len(v) > 0 for v in buf.values()):
                    resultado = calcular_ROSA_completo(
                        ang_tronco   = float(np.mean(buf["tronco"])),
                        ang_rodilla  = float(np.mean(buf["rodilla"])),
                        ang_codo     = float(np.mean(buf["codo"])),
                        desv_cuello  = float(np.mean(buf["cuello"])),
                        ang_muneca   = float(np.mean(buf["muneca"])),
                        horas_diarias= self.horas_diarias,
                    )
                    resultado["time"] = time.strftime("%Y-%m-%d %H:%M:%S")
                    self.cola_datos.put(resultado)
                    for k in self._buf:
                        self._buf[k].clear()
                last_eval = time.time()

        cap.release()
        landmarker.close()
        print("[hilo] Cámara cerrada.")

# ─────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────
def nivel_accion(score):
    k = min(max(score,1),5)
    return NIVELES[k]


class PanelROSA(tk.Tk):
    def __init__(self, nombre, horas_diarias, cam_idx, ruta_xlsx):
        super().__init__()
        self.title(f"Evaluador ROSA — NTP 1173  |  {nombre}")
        self.geometry("1150x660")
        self.configure(bg=C["bg_deep"])
        self.resizable(True, True)

        self.cola_datos   = queue.Queue()
        self.cola_frames  = queue.Queue(maxsize=2)
        self.cola_angulos = queue.Queue(maxsize=2)   # ángulos tiempo real
        self.registros    = []
        self.ruta_xlsx    = ruta_xlsx
        self.num_registro = 0
        self._ultimo_guardado = time.time()

        crear_excel_nuevo(self.ruta_xlsx)
        self._setup_ui()

        self.hilo = HiloCamara(self.cola_datos, self.cola_frames,
                               self.cola_angulos,
                               horas_diarias, cam_idx)
        self.hilo.start()
        self._update_loop()

    def _setup_ui(self):
        left = tk.Frame(self, bg=C["bg_deep"])
        left.pack(side="left", fill="y", padx=(16,8), pady=16)

        self.canvas = tk.Canvas(left, width=640, height=480,
                                bg="black", highlightthickness=1,
                                highlightbackground=C["accent2"])
        self.canvas.pack()
        self.img_id = self.canvas.create_image(0,0,anchor="nw")

        status_bar = tk.Frame(left, bg=C["bg_panel"], height=28)
        status_bar.pack(fill="x", pady=(4,0))
        self.lbl_fps = tk.Label(status_bar, text="Inicializando...",
                                font=("Consolas",8), bg=C["bg_panel"],
                                fg=C["text_lo"])
        self.lbl_fps.pack(side="left", padx=6)
        self.lbl_xlsx_status = tk.Label(status_bar, text="",
                                        font=("Consolas",8),
                                        bg=C["bg_panel"], fg=C["ok"])
        self.lbl_xlsx_status.pack(side="right", padx=6)

        right = tk.Frame(self, bg=C["bg_deep"])
        right.pack(side="right", fill="both", expand=True, padx=(0,16), pady=16)

        score_frame = tk.Frame(right, bg=C["bg_card"],
                               highlightthickness=1,
                               highlightbackground=C["bg_border"])
        score_frame.pack(fill="x", pady=(0,8))
        tk.Label(score_frame, text="PUNTUACIÓN ROSA FINAL",
                 font=("Consolas",9,"bold"), bg=C["bg_card"],
                 fg=C["text_lo"]).pack(pady=(10,0))
        self.lbl_score = tk.Label(score_frame, text="--",
                                  font=("Consolas",72,"bold"),
                                  bg=C["bg_card"], fg=C["accent"])
        self.lbl_score.pack()
        self.lbl_nivel = tk.Label(score_frame, text="EN ESPERA",
                                  font=("Consolas",12,"bold"),
                                  bg=C["bg_card"], fg=C["text_mid"])
        self.lbl_nivel.pack(pady=(0,10))

        tbl_frame = tk.Frame(right, bg=C["bg_card"],
                             highlightthickness=1,
                             highlightbackground=C["bg_border"])
        tbl_frame.pack(fill="both", expand=True)
        tk.Label(tbl_frame, text="DESGLOSE NTP 1173",
                 font=("Consolas",9,"bold"), bg=C["bg_card"],
                 fg=C["accent2"]).pack(pady=(8,4))

        grid = tk.Frame(tbl_frame, bg=C["bg_card"])
        grid.pack(padx=10, pady=4, fill="x")

        campos = [
            ("SILLA",           None,          True),
            ("A1 Altura",       "lbl_A1",      False),
            ("A2 Profundidad",  "lbl_A2",      False),
            ("A3 Reposabrazos", "lbl_A3",      False),
            ("A4 Respaldo",     "lbl_A4",      False),
            ("Tabla A",         "lbl_tA",      False),
            ("Factor Tiempo",   "lbl_ft",      False),
            ("Total Silla",     "lbl_tsilla",  False),
            ("",                None,          False),
            ("PERIFÉRICOS",     None,          True),
            ("B1 Teléfono",     "lbl_B1",      False),
            ("B2 Pantalla",     "lbl_B2",      False),
            ("Tabla B",         "lbl_tB",      False),
            ("C1 Ratón",        "lbl_C1",      False),
            ("C2 Teclado",      "lbl_C2",      False),
            ("Tabla C",         "lbl_tC",      False),
            ("Tabla D",         "lbl_tD",      False),
            ("",                None,          False),
            ("ÁNGULOS (°)",     None,          True),
            ("Tronco",          "lbl_at",      False),
            ("Rodilla",         "lbl_ar",      False),
            ("Codo",            "lbl_ac",      False),
            ("Cuello desv.",    "lbl_dc",      False),
            ("Muñeca",          "lbl_am",      False),
        ]

        for i, (nombre, attr, es_titulo) in enumerate(campos):
            if nombre == "":
                tk.Label(grid, text="", bg=C["bg_card"], height=1).grid(row=i, column=0)
                continue
            if es_titulo:
                tk.Label(grid, text=f"── {nombre} ──",
                         font=("Consolas",9,"bold"),
                         bg=C["bg_card"], fg=C["accent"],
                         anchor="w").grid(row=i, column=0, columnspan=2,
                                          sticky="w", pady=2)
            else:
                tk.Label(grid, text=nombre, font=("Consolas",9),
                         bg=C["bg_card"], fg=C["text_mid"],
                         anchor="w", width=18).grid(row=i, column=0, sticky="w")
                lbl = tk.Label(grid, text="—", font=("Consolas",9,"bold"),
                               bg=C["bg_card"], fg=C["text_hi"],
                               anchor="e", width=8)
                lbl.grid(row=i, column=1, sticky="e")
                if attr:
                    setattr(self, attr, lbl)
        grid.columnconfigure(0, weight=1)

        btn_frame = tk.Frame(right, bg=C["bg_deep"])
        btn_frame.pack(fill="x", pady=(8,0))
        tk.Button(btn_frame, text="GUARDAR EXCEL AHORA",
                  font=("Consolas",9,"bold"),
                  bg=C["accent2"], fg="white",
                  relief="flat", cursor="hand2",
                  command=self._guardar_xlsx_manual
                  ).pack(fill="x", ipady=6)
        tk.Label(btn_frame,
                 text=f"Auto-guardado cada {GUARDADO_SEG}s  |  {self.ruta_xlsx}",
                 font=("Consolas",7), bg=C["bg_deep"],
                 fg=C["text_lo"]).pack(pady=(2,0))

    def _actualizar_labels(self, d):
        self.lbl_score.config(text=str(d["rosa"]))
        txt, color = nivel_accion(d["rosa"])
        self.lbl_nivel.config(text=txt, fg=color)
        self.lbl_score.config(fg=color)
        def u(attr, val):
            lbl = getattr(self, attr, None)
            if lbl: lbl.config(text=str(val))
        u("lbl_A1", d["A1"]); u("lbl_A2", d["A2"])
        u("lbl_A3", d["A3"]); u("lbl_A4", d["A4"])
        u("lbl_tA", d["tabla_A"])
        u("lbl_ft", f"{d['factor_tiempo']:+d}")
        u("lbl_tsilla", d["total_silla"])
        u("lbl_B1", d["B1"]); u("lbl_B2", d["B2"])
        u("lbl_tB", d["tabla_B"])
        u("lbl_C1", d["C1"]); u("lbl_C2", d["C2"])
        u("lbl_tC", d["tabla_C"]); u("lbl_tD", d["tabla_D"])
        u("lbl_at", f"{d['ang_tronco']:+.1f}")
        u("lbl_ar", f"{d['ang_rodilla']:.1f}")
        u("lbl_ac", f"{d['ang_codo']:.1f}")
        u("lbl_dc", f"{d['desv_cuello']:+.1f}")
        u("lbl_am", f"{d['ang_muneca']:+.1f}")

    def _update_loop(self):
        try:
            frame_rgb = self.cola_frames.get_nowait()
            img = Image.fromarray(frame_rgb).resize((640,480))
            imgtk = ImageTk.PhotoImage(image=img)
            self.canvas.itemconfig(self.img_id, image=imgtk)
            self.canvas._imgtk = imgtk
            self.lbl_fps.config(text="Camara activa OK")
        except queue.Empty:
            pass

        # ── Ángulos en tiempo real (cada frame) ──
        try:
            ang = self.cola_angulos.get_nowait()
            def u(attr, val):
                lbl = getattr(self, attr, None)
                if lbl: lbl.config(text=str(val))
            u("lbl_at", f"{ang['at']:+.1f}")
            u("lbl_ar", f"{ang['ar']:.1f}")
            u("lbl_ac", f"{ang['ac']:.1f}")
            u("lbl_dc", f"{ang['dc']:+.1f}")
            u("lbl_am", f"{ang['am']:+.1f}")

            # Colorear según si está en rango óptimo
            def set_color(attr, val, ok_min, ok_max):
                lbl = getattr(self, attr, None)
                if lbl:
                    color = C["ok"] if ok_min <= abs(val) <= ok_max else C["danger"]
                    lbl.config(fg=color)
            set_color("lbl_at", ang["at"],  0,  15)
            set_color("lbl_ar", ang["ar"], 85, 100)
            set_color("lbl_ac", ang["ac"], 80, 110)
            set_color("lbl_dc", ang["dc"],  0,  10)
            set_color("lbl_am", ang["am"],  0,  15)
        except queue.Empty:
            pass

        # ── Dato de evaluación ROSA completo (cada 10s) ──
        try:
            dato = self.cola_datos.get_nowait()
            if "error" in dato:
                messagebox.showerror("Error", dato["error"])
                self.destroy()
                return
            self.registros.append(dato)
            self.num_registro += 1
            self._actualizar_labels(dato)
            if dato["rosa"] >= NIVEL_ACCION:
                self._alertar()
        except queue.Empty:
            pass

        if (self.registros and
                time.time() - self._ultimo_guardado >= GUARDADO_SEG):
            self._guardar_nuevo_registro()

        self.after(30, self._update_loop)

    def _guardar_nuevo_registro(self):
        if not self.registros: return
        try:
            agregar_registro_excel(self.ruta_xlsx,
                                   self.registros[-1], self.num_registro)
            self._ultimo_guardado = time.time()
            self.lbl_xlsx_status.config(
                text=f"Guardado {time.strftime('%H:%M:%S')}", fg=C["ok"])
        except Exception as e:
            self.lbl_xlsx_status.config(
                text=f"Error guardado: {e}", fg=C["danger"])

    def _guardar_xlsx_manual(self):
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            title="Guardar evaluación ROSA",
            initialfile=f"ROSA_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        if not ruta: return
        try:
            crear_excel_nuevo(ruta)
            for i, r in enumerate(self.registros, 1):
                agregar_registro_excel(ruta, r, i)
            messagebox.showinfo("Guardado", f"Excel guardado en:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _alertar(self):
        if WINSOUND_OK:
            threading.Thread(target=lambda: winsound.Beep(1000,500),
                             daemon=True).start()

    def on_close(self):
        self.hilo.activo = False
        self.destroy()

# ─────────────────────────────────────────────────────────────
# DIÁLOGO DE INICIO
# ─────────────────────────────────────────────────────────────
def detectar_camaras():
    camaras = []
    for i in range(MAX_CAMARAS_SCAN):
        c = cv2.VideoCapture(i, cv2.CAP_DSHOW)
        if c.isOpened():
            camaras.append((i, f"Camara {i}"))
            c.release()
    return camaras if camaras else [(0,"Default")]


class DialogoInicio(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ROSA NTP 1173 — Configuracion")
        self.geometry("420x420")
        self.configure(bg=C["bg_deep"])
        self.resizable(False, False)
        self.resultado = None
        self._build()

    def _build(self):
        tk.Label(self, text="EVALUADOR ROSA",
                 font=("Consolas",18,"bold"),
                 bg=C["bg_deep"], fg=C["accent"]).pack(pady=(24,2))
        tk.Label(self, text="NTP 1173 · INSST 2022 · MediaPipe 0.10",
                 font=("Consolas",9),
                 bg=C["bg_deep"], fg=C["text_lo"]).pack(pady=(0,4))

        # Estado descarga modelo
        self.lbl_modelo = tk.Label(self,
                 text="Verificando modelo...",
                 font=("Consolas",8), bg=C["bg_deep"], fg=C["warn"])
        self.lbl_modelo.pack(pady=(0,12))

        frame = tk.Frame(self, bg=C["bg_card"], padx=20, pady=16)
        frame.pack(fill="x", padx=30)

        def row(lbl_txt, widget_fn):
            tk.Label(frame, text=lbl_txt, font=("Consolas",9),
                     bg=C["bg_card"], fg=C["text_mid"],
                     anchor="w").pack(fill="x", pady=(4,0))
            w = widget_fn(frame)
            w.pack(fill="x", pady=(2,0))
            return w

        self.ent_nombre = row("Nombre del trabajador:", lambda p: tk.Entry(
            p, font=("Consolas",10), bg=C["bg_border"],
            fg=C["text_hi"], insertbackground="white", relief="flat"))
        self.ent_nombre.insert(0, "Trabajador 01")

        self.spin_horas = row("Horas diarias frente al equipo:",
                               lambda p: ttk.Spinbox(p, from_=0.5, to=12,
                               increment=0.5, width=8, font=("Consolas",10)))
        self.spin_horas.set("6")

        camaras = detectar_camaras()
        self.combo_cam = row("Camara:", lambda p: ttk.Combobox(
            p, values=[f"{i}: {n}" for i,n in camaras],
            state="readonly", font=("Consolas",10)))
        self.combo_cam.current(0)
        self._camaras = camaras

        tk.Label(self, text="", bg=C["bg_deep"]).pack()

        self.btn_iniciar = tk.Button(self, text="▶  INICIAR EVALUACION",
                  font=("Consolas",11,"bold"),
                  bg=C["accent"], fg="white", relief="flat",
                  cursor="hand2", command=self._ok, state="disabled")
        self.btn_iniciar.pack(fill="x", padx=30, ipady=8)

        # Descargar modelo en background
        threading.Thread(target=self._verificar_modelo, daemon=True).start()

    def _verificar_modelo(self):
        def status(msg):
            self.lbl_modelo.config(text=msg)
        ok = descargar_modelo(status)
        if ok:
            self.lbl_modelo.config(
                text="Modelo listo OK", fg=C["ok"])
            self.btn_iniciar.config(state="normal")
        else:
            self.lbl_modelo.config(
                text="ERROR: no se pudo descargar el modelo", fg=C["danger"])

    def _ok(self):
        nombre  = self.ent_nombre.get().strip() or "Trabajador"
        horas   = float(self.spin_horas.get())
        idx_cam = self._camaras[self.combo_cam.current()][0]
        ruta_xlsx = os.path.join(
            os.path.expanduser("~"), "Desktop",
            f"ROSA_{nombre.replace(' ','_')}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        self.resultado = (nombre, horas, idx_cam, ruta_xlsx)
        self.destroy()

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("[main] Iniciando Evaluador ROSA NTP-1173 v4.1")
    print(f"[main] Modelo: {MODEL_PATH}")

    dialogo = DialogoInicio()
    dialogo.mainloop()

    if dialogo.resultado is None:
        print("[main] Cancelado por el usuario.")
        sys.exit(0)

    nombre, horas, cam_idx, ruta_xlsx = dialogo.resultado
    print(f"[main] Trabajador: {nombre} | Horas: {horas} | Cam: {cam_idx}")
    print(f"[main] Excel: {ruta_xlsx}")

    app = PanelROSA(nombre, horas, cam_idx, ruta_xlsx)
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
    print("[main] Aplicación cerrada.")