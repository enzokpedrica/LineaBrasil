import cv2
import numpy as np
import matplotlib.pyplot as plt

# === CONFIGURAÇÕES DO PROJETO ===
pixels_por_cm = 58.3
x_esperado_cm = 9.86  
y_esperado_cm = 3.76  
raio_esperado_cm = 0.34  
tolerancia_cm = 1.5  
tolerancia_raio_cm = 0.7  

# Conversão para pixels
x_esperado = x_esperado_cm * pixels_por_cm
y_esperado = y_esperado_cm * pixels_por_cm
raio_esperado = raio_esperado_cm * pixels_por_cm
tolerancia_pos = tolerancia_cm * pixels_por_cm
tolerancia_raio = tolerancia_raio_cm * pixels_por_cm

# === CARREGAR IMAGEM ===
caminho_imagem = "imagem.jpeg"  
imagem = cv2.imread(caminho_imagem)
saida = imagem.copy()

# Pré-processamento
gray = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
blur = cv2.GaussianBlur(gray, (11, 11), 2.5)
edges = cv2.Canny(blur, 30, 100)

# === DETECÇÃO DE FURO ===
circles = cv2.HoughCircles(
    edges,
    cv2.HOUGH_GRADIENT,
    dp=1.2,
    minDist=50,
    param1=100,
    param2=30,
    minRadius=15,
    maxRadius=50
)

# === ANÁLISE ===
resultado = "Nenhum furo detectado"
if circles is not None:
    circles = np.round(circles[0, :]).astype("int")
    for (x, y, r) in circles:
        print(f"Detectado: x={x}, y={y}, r={r}")
        pos_ok = (abs(x - x_esperado) <= tolerancia_pos and
                  abs(y - y_esperado) <= tolerancia_pos)
        raio_ok = abs(r - raio_esperado) <= tolerancia_raio

        cor = (0, 255, 0) if pos_ok and raio_ok else (0, 0, 255)
        texto = "Furo OK" if pos_ok and raio_ok else "Furo Incorreto"
        cv2.circle(saida, (x, y), r, cor, 2)
        cv2.circle(saida, (x, y), 2, (0, 255, 255), 3)
        cv2.putText(saida, texto, (x - 100, y - 40),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, cor, 2)

        resultado = texto

# === EXIBIR RESULTADO ===
saida_rgb = cv2.cvtColor(saida, cv2.COLOR_BGR2RGB)
plt.figure(figsize=(10, 8))
plt.imshow(saida_rgb)
plt.title(resultado)
plt.axis("off")
plt.show()