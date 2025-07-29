#!/usr/bin/env bash
# build.sh - Otimizado para Render.com

set -o errexit  # Sai imediatamente em caso de erro

echo "--> Atualizando pacotes do sistema"
apt-get update -qq
apt-get install -y --no-install-recommends \
    python3-dev \
    build-essential \
    libpq-dev  # Apenas se usar PostgreSQL

echo "--> Configurando ambiente Python"
python -m pip install --upgrade pip setuptools wheel

echo "--> Instalando dependências Python"
pip install --no-cache-dir -r requirements.txt

echo "--> Limpando cache para reduzir espaço"
apt-get autoremove -y
apt-get clean -y
rm -rf /var/lib/apt/lists/*
