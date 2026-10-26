FROM python:3.12-slim

WORKDIR /app

# onnxruntime and headless OpenCV need these system libraries.
RUN apt-get update \
    && apt-get install -y --no-install-recommends libgomp1 libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Fetch the background-removal model once, at build time, so the first batch
# does not wait for it. If the network is unavailable here, the worker simply
# downloads it the first time it is needed.
RUN python -c "from rembg import new_session; new_session('u2netp')" || true

COPY . .

CMD ["python", "-u", "worker.py"]
