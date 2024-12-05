# Use a imagem base Python slim para economizar espaço
FROM python:3.9-slim

# Configurar o diretório de trabalho dentro do contêiner
WORKDIR /app

# Copiar todos os arquivos para o diretório de trabalho do contêiner
COPY . .

# Instalar dependências do projeto
RUN pip install --no-cache-dir -r requirements.txt

# Certifique-se de que o .env será reconhecido no contêiner
ENV PYTHONUNBUFFERED=1

# Comando para executar o script
CMD ["python", "integracao.py"]
