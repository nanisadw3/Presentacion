#!/bin/bash
set -e

# 1. Iniciar Xvfb (Servidor X virtual en memoria)
echo "Iniciando Xvfb en display :99 con resolución ${RESOLUTION}..."
Xvfb :99 -screen 0 ${RESOLUTION} &
sleep 2

# 2. Iniciar Openbox (Gestor de ventanas para mover/redimensionar la ventana de la app)
echo "Iniciando Openbox..."
openbox-session &
sleep 1

# 3. Iniciar la aplicación CustomTkinter en segundo plano
echo "Iniciando la aplicación de Presentación..."
python app.py &
sleep 2

# 4. Iniciar x11vnc (Servidor VNC para transmitir el display virtual)
echo "Iniciando x11vnc..."
x11vnc -display :99 -nopw -forever -shared -bg
sleep 1

# 5. Iniciar websockify para enlazar noVNC en el puerto 8080
echo "Iniciando noVNC Proxy en el puerto 8080..."
websockify --web /usr/share/novnc 8080 localhost:5900
