import cv2
import mediapipe as mp
import numpy as np
import time
import csv
import os

# ==============================
# DATOS INICIALES
# ==============================
nombre = input("Nombre del trabajador: ")
id_trabajador = input("ID trabajador: ")
archivo_csv = f"evaluacion_ROSA_{id_trabajador}.csv"
puntos_tiempo = 2 # Suponemos jornada completa (+1) y (+1) si la silla no es ajustable

if not os.path.exists(archivo_csv):
    with open(archivo_csv, mode='w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Timestamp', 'Punt_Tronco', 'Punt_Cuello', 'Punt_Rodilla', 'Punt_Final_ROSA', 'Estado'])

# Tabla A simplificada (Cruce de puntuaciones)
TABLA_A = [
    [2,2,3,4,5,6,7,8], [2,2,3,4,5,6,7,8], [3,3,3,4,5,6,7,8],
    [4,4,4,4,5,6,7,8], [5,5,5,5,6,7,8,9], [6,6,6,7,7,8,8,9],
    [7,7,7,8,8,9,9,9]
]

# ==============================
# CONFIGURACIÓN MEDIAPIPE
# ==============================
mp_pose = mp.solutions.pose
mp_draw = mp.solutions.drawing_utils
pose = mp_pose.Pose(min_detection_confidence=0.5, min_tracking_confidence=0.5)

def calcular_angulo(a, b, c):
    a, b, c = np.array(a), np.array(b), np.array(c)
    ba, bc = a - b, c - b
    n_ba, n_bc = np.linalg.norm(ba), np.linalg.norm(bc)
    if n_ba == 0 or n_bc == 0: return 0
    return np.degrees(np.arccos(np.clip(np.dot(ba, bc) / (n_ba * n_bc), -1.0, 1.0)))

# ==============================
# VARIABLES DE CONTROL
# ==============================
inicio_ventana = time.time()
buffer_datos = []
tiempo_riesgo_consecutivo = 0 
alerta_activa = False
p_rosa_actual = 1

cap = cv2.VideoCapture(0)

# Configurar Resolucion a 720p (Ideal para 60fps en C922)
cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

# Forzar FPS (La C922 lo soporta a 720p)
# cap.set(cv2.CAP_PROP_FPS, 60)

# Desactivar Autoenfoque (IMPORTANTE)
# El autoenfoque constante puede hacer que MediaPipe pierda los landmarks
# cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)

while cap.isOpened():
    ret, frame = cap.read()
    if not ret: break
    
    frame = cv2.flip(frame, 1)
    h, w, _ = frame.shape
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = pose.process(rgb)
    
    angulos = {"tronco": 0, "cuello": 0, "rodilla": 0}

    if results.pose_landmarks:
        lm = results.pose_landmarks.landmark
        ls, rs = lm[mp_pose.PoseLandmark.LEFT_SHOULDER], lm[mp_pose.PoseLandmark.RIGHT_SHOULDER]
        
        if abs(ls.visibility - rs.visibility) > 0.1 or abs(ls.x - rs.x) < 0.15:
            lado = "LEFT" if ls.visibility > rs.visibility else "RIGHT"
            
            oreja = [lm[getattr(mp_pose.PoseLandmark, f"{lado}_EAR")].x * w, lm[getattr(mp_pose.PoseLandmark, f"{lado}_EAR")].y * h]
            hombro = [lm[getattr(mp_pose.PoseLandmark, f"{lado}_SHOULDER")].x * w, lm[getattr(mp_pose.PoseLandmark, f"{lado}_SHOULDER")].y * h]
            cadera = [lm[getattr(mp_pose.PoseLandmark, f"{lado}_HIP")].x * w, lm[getattr(mp_pose.PoseLandmark, f"{lado}_HIP")].y * h]
            rodilla = [lm[getattr(mp_pose.PoseLandmark, f"{lado}_KNEE")].x * w, lm[getattr(mp_pose.PoseLandmark, f"{lado}_KNEE")].y * h]
            tobillo = [lm[getattr(mp_pose.PoseLandmark, f"{lado}_ANKLE")].x * w, lm[getattr(mp_pose.PoseLandmark, f"{lado}_ANKLE")].y * h]

            angulos["tronco"] = calcular_angulo(hombro, cadera, rodilla)
            angulos["cuello"] = calcular_angulo(oreja, hombro, [hombro[0], 0])
            angulos["rodilla"] = calcular_angulo(cadera, rodilla, tobillo)
            
            buffer_datos.append(list(angulos.values()))
        
        mp_draw.draw_landmarks(frame, results.pose_landmarks, mp_pose.POSE_CONNECTIONS)

    # ==========================================
    # LÓGICA DE EVALUACIÓN ROSA (CADA 5 SEG)
    # ==========================================
    ahora = time.time()
    if ahora - inicio_ventana > 5:
        if buffer_datos:
            t, c, r = np.mean(buffer_datos, axis=0)
            
            # Puntuación individual (1 punto es ideal, >1 es desviación)
            p_tronco = 1 if (95 <= t <= 110) else 2
            p_cuello = 1 if (c <= 20) else (2 if c <= 30 else 3)
            p_rodilla = 1 if (90 <= r <= 110) else 2
            
            # Puntuación Silla (Cruce Tabla A)
            idx_a = min(p_rodilla - 1, 6)
            idx_r = min(p_tronco - 1, 7)
            p_silla = TABLA_A[idx_a][idx_r] + puntos_tiempo
            
            # Puntuación Final (Basada en el máximo riesgo detectado)
            p_pantalla = p_cuello + puntos_tiempo
            p_rosa_actual = max(p_silla, p_pantalla)
            
            # --- CRITERIO DE ALERTA ---
            # Si la puntuación es 5 o más, se considera Riesgo Alto (Nivel de Actuación 2)
            if p_rosa_actual >= 5:
                tiempo_riesgo_consecutivo += 5
            else:
                tiempo_riesgo_consecutivo = 0 # Reset si baja de 5

            alerta_activa = True if tiempo_riesgo_consecutivo >= 30 else False

            with open(archivo_csv, mode='a', newline='') as f:
                writer = csv.writer(f)
                writer.writerow([time.strftime('%H:%M:%S'), p_tronco, p_cuello, p_rodilla, p_rosa_actual, "ALERTA" if alerta_activa else "NORMAL"])

            buffer_datos = []
        inicio_ventana = ahora

    # ==============================
    # INTERFAZ GRÁFICA
    # ==============================
    if alerta_activa:
        cv2.rectangle(frame, (0,0), (w,h), (0,0,255), 25)
        cv2.putText(frame, "!!! ALERTA ROSA: RIESGO ALTO !!!", (w//6, h//2), 2, 1.3, (0,0,255), 4)

    # Panel lateral
    cv2.rectangle(frame, (10, 10), (450, 200), (0,0,0), -1)
    color_puntos = (0,255,0) if p_rosa_actual < 5 else (0,0,255)
    cv2.putText(frame, f"PUNTUACION ROSA: {p_rosa_actual}", (20, 50), 1, 1.8, color_puntos, 2)
    cv2.putText(frame, f"Tronco: {int(angulos['tronco'])} | Cuello: {int(angulos['cuello'])}", (20, 100), 1, 1, (255,255,255), 1)
    cv2.putText(frame, f"Tiempo en Riesgo: {tiempo_riesgo_consecutivo}s / 30s", (20, 160), 1, 1.4, (0,255,255), 2)

    cv2.imshow("Sistema Ergonomico ROSA", frame)
    if cv2.waitKey(1) & 0xFF == 27: break

cap.release()
cv2.destroyAllWindows()