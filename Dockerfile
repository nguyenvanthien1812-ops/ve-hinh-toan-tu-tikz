FROM ubuntu:22.04

ENV DEBIAN_frontend=noninteractive
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Cài TeX + Poppler + Python
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip \
    texlive-latex-base \
    texlive-latex-extra \
    texlive-science \
    texlive-pictures \
    poppler-utils \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]
