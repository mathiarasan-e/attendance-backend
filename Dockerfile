FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY . .

# Catalyst AppSail injects the port to listen on via this variable.
# It defaults to 9000 when running the image locally.
ENV X_ZOHO_CATALYST_LISTEN_PORT=9000
EXPOSE 9000

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${X_ZOHO_CATALYST_LISTEN_PORT:-9000}"]
