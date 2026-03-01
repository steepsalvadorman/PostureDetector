import cv2
import math as m
import mediapipe as mp
import time
import numpy as np

# --- 1. MATRICES DE CRUCE OFICIALES (MÉTODO ROSA) ---

# Tabla A: Silla (Altura+Profundidad vs Reposabrazos+Respaldo)
TABLA_A = np.array([
    [2, 2, 3, 4, 5, 6, 7, 8],
    [2, 2, 3, 4, 5, 6, 7, 8],
    [3, 3, 3, 4, 5, 6, 7, 8],
    [4, 4, 4, 4, 5, 6, 7, 8],
    [5, 5, 5, 5, 5, 6, 7, 8],
    [6, 6, 6, 6, 6, 6, 7, 8],
    [7, 7, 7, 7, 7, 7, 8, 9],
    [8, 8, 8, 8, 8, 8, 9, 9]
])

# Tabla B: Pantalla y Teléfono
TABLA_B = np.array([
    [1, 2, 3, 4, 5, 6],
    [2, 2, 3, 4, 5, 6],
    [2, 3, 3, 4, 6, 7],
    [2, 3, 4, 5, 6, 8],
    [3, 4, 4, 5, 6, 8],
    [4, 5, 5, 6, 7, 8]
])

# Tabla C: Mouse y Teclado
TABLA_C = np.array([
    [1, 1, 1, 2, 3, 4, 5, 6],
    [1, 1, 2, 3, 4, 5, 6, 7],
    [1, 2, 2, 3, 4, 5, 6, 7],
    [2, 3, 3, 3, 5, 6, 7, 8],
    [3, 4, 4, 5, 5, 6, 7, 8],
    [4, 5, 5, 6, 6, 7, 8, 9],
    [5, 6, 6, 7, 7, 8, 8, 9]
])

# Tabla D: Periféricos (Resultado B vs Resultado C)
TABLA_D = np.array([
    [1, 2, 3, 4, 5, 6, 7, 8, 9],
    [2, 2, 3, 4, 5, 6, 7, 8, 9],
    [3, 3, 3, 4, 5, 6, 7, 8, 9],
    [4, 4, 4, 4, 5, 6, 7, 8, 9],
    [5, 5, 5, 5, 5, 6, 7, 8, 9],
    [6, 6, 6, 6, 6, 6, 7, 8, 9],
    [7, 7, 7, 7, 7, 7, 7, 8, 9]
])

# Tabla E: PUNTUACIÓN FINAL (Silla vs Periféricos)
TABLA_E = np.identity(11) # Simplificación para el ejemplo, representa el cruce final

# --- 2. CONFIGURACIÓN INICIAL (AUDITORÍA) ---
# Estos valores se obtendrían de una entrevista previa o checklist
uso_diario_horas = 5 # >4 horas suma +1 según Tabla 7
silla_no_ajustable = True # +1 punto
mouse_pequeno = False
teclado_muy_alto = False

# --- 3. FUNCIONES DE APOYO ---

def calcular_angulo(p1, p2, p3):
    try:
        rad = m.atan2(p3[1]-p2[1], p3[0]-p2[0]) - m.atan2(p1[1]-p2[1], p1[0]-p2[0])
        ang = abs(rad * 180.0 / m.pi)
        return 360 - ang if ang > 180 else ang
    except: return 0

def obtener_nivel_actuacion(puntaje):
    if puntaje <= 1: return "INAPRECIABLE", 0, (127, 255, 0)
    if 2 <= puntaje <= 4: return "MEJORABLE", 1, (0, 255, 255)
    if puntaje == 5: return "ALTO", 2, (0, 165, 255)
    if 6 <= puntaje <= 8: return "MUY ALTO", 3, (50, 50, 255)
    return "EXTREMO", 4, (0, 0, 255)

# --- 4. PROCESAMIENTO PRINCIPAL ---

mp_pose = mp.solutions.pose
pose = mp_pose.Pose(model_complexity=2, min_detection_confidence=0.8)
cap = cv2.VideoCapture(0)

while cap.isOpened():
    ret, frame = cap.read()
    if not ret: break
    
    h, w = frame.shape[:2]
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    res = pose.process(rgb)

    if res.pose_landmarks:
        lm = res.pose_landmarks.landmark
        idx = mp_pose.PoseLandmark

        # Selección de perfil
        lado = "LEFT" if lm[idx.LEFT_SHOULDER].visibility > lm[idx.RIGHT_SHOULDER].visibility else "RIGHT"
        
        # Coordenadas
        def get_p(part):
            p = lm[getattr(idx, f"{lado}_{part}")]
            return (int(p.x * w), int(p.y * h))

        p_oreja = (int(lm[getattr(idx, f"{lado}_EAR")].x * w), int(lm[getattr(idx, f"{lado}_EAR")].y * h))
        p_hombro, p_cadera, p_rodilla, p_tobillo = get_p("SHOULDER"), get_p("HIP"), get_p("KNEE"), get_p("ANKLE")

        # --- CÁLCULOS BIOMECÁNICOS ---
        ang_cuello = calcular_angulo(p_oreja, p_hombro, (p_hombro[0], 0))
        ang_tronco = calcular_angulo(p_hombro, p_cadera, (p_cadera[0], 0))
        ang_rodilla = calcular_angulo(p_cadera, p_rodilla, p_tobillo)

        # --- APLICACIÓN DE PUNTUACIONES ROSA ---
        # Silla
        p_altura = (2 if ang_rodilla < 90 or ang_rodilla > 95 else 1) + (1 if silla_no_ajustable else 0)
        p_respaldo = (2 if ang_tronco < 95 or ang_tronco > 110 else 1)
        
        # Ajuste por tiempo (Tabla 7)
        mod_tiempo = 1 if uso_diario_horas > 4 else 0

        # Cruce Tabla A (Simplificado para el demo)
        val_tabla_a = TABLA_A[min(p_altura, 7)][min(p_respaldo + 1, 7)] + mod_tiempo
        
        # Puntuación Final Simulada (Cruce de Silla y Periféricos)
        # En un sistema completo, aquí se calcularían Pantalla, Mouse y Teclado
        puntuacion_final = min(val_tabla_a + 1, 10) 
        
        txt_riesgo, nivel, color = obtener_nivel_actuacion(puntuacion_final)

        # --- INTERFAZ ---
        # Dibujar esqueleto y HUD
        cv2.rectangle(frame, (0,0), (380, 180), (30,30,30), -1)
        cv2.putText(frame, f"PUNTAJE ROSA: {puntuacion_final}", (20, 40), 1, 2, color, 3)
        cv2.putText(frame, f"RIESGO: {txt_riesgo}", (20, 80), 1, 1.5, (255,255,255), 2)
        cv2.putText(frame, f"NIVEL DE ACTUACION: {nivel}", (20, 120), 1, 1.2, color, 2)
        
        # Feedback visual de ángulos
        cv2.putText(frame, f"Cuello: {int(ang_cuello)} deg", (20, 150), 1, 1, (255,255,0), 1)
        cv2.line(frame, p_hombro, p_oreja, color, 4)
        cv2.line(frame, p_hombro, p_cadera, color, 4)
        cv2.line(frame, p_cadera, p_rodilla, color, 4)

    cv2.imshow("Sistema ROSA Profesional - Ergonautas UPV", frame)
    if cv2.waitKey(1) & 0xFF == ord('q'): break

cap.release()
cv2.destroyAllWindows()