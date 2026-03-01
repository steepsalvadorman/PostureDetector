import cv2
import mediapipe as mp
import math as m
import time
import os
import threading
import pandas as pd
from datetime import datetime
from ultralytics import YOLO

# ==========================================
# 1. CAPA DE STREAMING ROBUSTO (Threading)
# ==========================================
class EzvizStream:
    """Clase para leer la cámara EZVIZ en un hilo separado y evitar lag acumulado."""
    def __init__(self, rtsp_url):
        self.cap = cv2.VideoCapture(rtsp_url)
        self.cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)  # Buffer mínimo para tiempo real
        self.ret = False
        self.frame = None
        self.stopped = False

    def start(self):
        threading.Thread(target=self.update, args=(), daemon=True).start()
        return self

    def update(self):
        while not self.stopped:
            if not self.cap.isOpened():
                self.ret = False
            else:
                (self.ret, self.frame) = self.cap.read()
            if not self.ret:
                time.sleep(0.01)

    def read(self):
        return self.frame

    def stop(self):
        self.stopped = True
        self.cap.release()

# ==========================================
# 2. CAPA DE PERSISTENCIA (Data Logging)
# ==========================================
class ErgoPersistence:
    def __init__(self, filename="log_ergonomico.csv"):
        self.filename = filename
        if not os.path.exists(self.filename):
            df = pd.DataFrame(columns=[
                "Fecha", "Hora", "ROSA_Score", "Nivel_Riesgo", 
                "Ang_Cuello", "Ang_Tronco", "Seg_Mala_Postura"
            ])
            df.to_csv(self.filename, index=False)

    def registrar(self, score, n_ang, t_ang, bad_time):
        if score <= 3: riesgo = "Bajo (Normal)"
        elif score <= 5: riesgo = "Medio (Precaucion)"
        else: riesgo = "Alto (Peligro: Intervencion)"

        nuevo_registro = {
            "Fecha": datetime.now().strftime("%Y-%m-%d"),
            "Hora": datetime.now().strftime("%H:%M:%S"),
            "ROSA_Score": score,
            "Nivel_Riesgo": riesgo,
            "Ang_Cuello": round(n_ang, 2),
            "Ang_Tronco": round(t_ang, 2),
            "Seg_Mala_Postura": int(bad_time)
        }
        pd.DataFrame([nuevo_registro]).to_csv(self.filename, mode='a', header=False, index=False)

# ==========================================
# 3. MOTOR DE CÁLCULOS BIOMECÁNICOS
# ==========================================
def calculate_angle(p1, p2, p3):
    try:
        radians = m.atan2(p3[1] - p2[1], p3[0] - p2[0]) - m.atan2(p1[1] - p2[1], p1[0] - p2[0])
        angle = abs(radians * 180.0 / m.pi)
        if angle > 180.0: angle = 360 - angle
        return angle
    except: return 0

def get_rosa_logic(n_ang, t_ang, k_ang):
    sn = 1 if n_ang <= 20 else 2 if n_ang <= 45 else 3
    st = 1 if 95 <= t_ang <= 110 else 2
    sk = 1 if 90 <= k_ang <= 105 else 2
    return sn, st, sk

# ==========================================
# 4. CONFIGURACIÓN E INICIO
# ==========================================
# Configuración de Modelos
model_yolo = YOLO('yolov8n.pt') 
mp_pose = mp.solutions.pose
mp_drawing = mp.solutions.drawing_utils
# Complexity 0 es vital para cámaras Wi-Fi/Batería para reducir latencia
pose = mp_pose.Pose(model_complexity=0, min_detection_confidence=0.6) 

# Paleta de colores
RED, GREEN, YELLOW, CYAN, WHITE = (50, 50, 255), (127, 255, 0), (0, 255, 255), (255, 255, 0), (255, 255, 255)

if __name__ == "__main__":
    # --- CONFIGURACIÓN DE TU CÁMARA EZVIZ ---
    # Reemplaza 'CODIGO' por el código de verificación en la etiqueta y 'IP' por la de tu red
    URL_RTSP = "rtsp://admin:CODIGO@192.168.1.XX:554/ch1/main"
    
    # Iniciar capturador en hilo separado
    cam = EzvizStream(URL_RTSP).start()
    db = ErgoPersistence()
    
    posture_start_time = time.time()
    last_save_time = time.time()
    frame_count = 0
    results_yolo = []

    print("Iniciando Analizador... Presiona 'q' para salir.")

    while True:
        frame = cam.read()
        if frame is None:
            continue
            
        h, w = frame.shape[:2]
        frame_count += 1

        # A. VISIÓN DE ENTORNO (Cada 30 frames para no saturar CPU)
        if frame_count % 30 == 0:
            results_yolo = model_yolo(frame, classes=[62, 63, 66], verbose=False)

        for r in results_yolo:
            for box in r.boxes:
                x1, y1, x2, y2 = map(int, box.xyxy[0])
                cv2.rectangle(frame, (x1, y1), (x2, y2), CYAN, 1)
                cv2.putText(frame, "EQUIPO", (x1, y1-5), 0, 0.4, CYAN, 1)

        # B. VISIÓN HUMANA (MediaPipe)
        # Procesamos en RGB
        results_pose = pose.process(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))

        if results_pose.pose_landmarks:
            mp_drawing.draw_landmarks(frame, results_pose.pose_landmarks, mp_pose.POSE_CONNECTIONS,
                mp_drawing.DrawingSpec(color=WHITE, thickness=1, circle_radius=1),
                mp_drawing.DrawingSpec(color=GREEN, thickness=2, circle_radius=1))

            try:
                lm = results_pose.pose_landmarks.landmark
                lm_p = mp_pose.PoseLandmark

                # Selección automática del lado más visible
                side = "L" if lm[lm_p.LEFT_SHOULDER].visibility > lm[lm_p.RIGHT_SHOULDER].visibility else "R"
                if side == "L":
                    ear, shldr, hip, knee = lm[lm_p.LEFT_EAR], lm[lm_p.LEFT_SHOULDER], lm[lm_p.LEFT_HIP], lm[lm_p.LEFT_KNEE]
                else:
                    ear, shldr, hip, knee = lm[lm_p.RIGHT_EAR], lm[lm_p.RIGHT_SHOULDER], lm[lm_p.RIGHT_HIP], lm[lm_p.RIGHT_KNEE]

                p_ear = (int(ear.x * w), int(ear.y * h))
                p_shldr = (int(shldr.x * w), int(shldr.y * h))
                p_hip = (int(hip.x * w), int(hip.y * h))
                p_knee = (int(knee.x * w), int(knee.y * h))

                # Cálculos ROSA (Ángulos contra la vertical)
                n_ang = calculate_angle(p_ear, p_shldr, (p_shldr[0], 0))
                t_ang = calculate_angle(p_shldr, p_hip, (p_hip[0], 0))
                k_ang = 95 # Valor por defecto si sentado

                sn, st, sk = get_rosa_logic(n_ang, t_ang, k_ang)
                
                # Gestión de tiempo de mala postura
                if (sn + st) > 2:
                    elapsed_time = time.time() - posture_start_time
                else:
                    posture_start_time = time.time()
                    elapsed_time = 0

                penalty = 1 if elapsed_time > 10 else 0
                rosa_final = min((sn + st + sk + penalty), 10)
                color = GREEN if rosa_final <= 3 else YELLOW if rosa_final <= 5 else RED

                # PERSISTENCIA (Cada 5 segundos)
                if time.time() - last_save_time > 5:
                    db.registrar(rosa_final, n_ang, t_ang, elapsed_time)
                    last_save_time = time.time()

                # INTERFAZ (HUD)
                cv2.rectangle(frame, (10, 10), (350, 150), (30,30,30), -1)
                cv2.putText(frame, f"ROSA SCORE: {rosa_final}/10", (20, 45), 0, 0.8, color, 2)
                cv2.putText(frame, f"RIESGO: {int(elapsed_time)}s", (20, 80), 0, 0.6, CYAN, 1)
                cv2.putText(frame, f"SIDE: {side} | CUELLO: {int(n_ang)}", (20, 110), 0, 0.5, WHITE, 1)
                cv2.putText(frame, f"TRONCO: {int(t_ang)}", (20, 135), 0, 0.5, WHITE, 1)

            except Exception: pass
        else:
            posture_start_time = time.time()

        cv2.imshow('ANALIZADOR ERGONOMICO - EZVIZ CB2', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'): break

    cam.stop()
    cv2.destroyAllWindows()