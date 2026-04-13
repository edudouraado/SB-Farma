import pyautogui
import time

print("Posicione o mouse sobre o botão em 5 segundos...")
time.sleep(5)
x, y = pyautogui.position()
print(f"As coordenadas são: X={x}, Y={y}")