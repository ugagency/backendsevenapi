FROM python:3.12-slim
ENV PYTHONUNBUFFERED=1
WORKDIR /srv
COPY os_api/requirements.txt /srv/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY os_api /srv/os_api
EXPOSE 8100
# Dentro do container escuta em todas as interfaces; no host a porta e publicada so em 127.0.0.1.
CMD ["uvicorn", "os_api.main:app_factory", "--factory", "--host", "0.0.0.0", "--port", "8100"]
