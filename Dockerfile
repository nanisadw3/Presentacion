FROM python:3.11-slim-bookworm

# Evitar prompts interactivos de apt
ENV DEBIAN_FRONTEND=noninteractive

# Instalar dependencias del sistema para la GUI, X11 y noVNC
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3-tk \
    xvfb \
    x11vnc \
    novnc \
    websockify \
    openbox \
    sqlite3 \
    libglib2.0-0 \
    && ln -s /usr/share/novnc/vnc.html /usr/share/novnc/index.html \
    && rm -rf /var/lib/apt/lists/*

# Configurar directorio de trabajo
WORKDIR /app

# Copiar dependencias e instalarlas
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el código del proyecto
COPY . .

# Copiar script de arranque y darle permisos de ejecución
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# Exponer el puerto 8080 para acceder por navegador
EXPOSE 8080

# Definir variables de entorno para Xvfb y display
ENV DISPLAY=:99
ENV RESOLUTION=1280x720x24

ENTRYPOINT ["/entrypoint.sh"]
