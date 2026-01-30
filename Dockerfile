FROM public.ecr.aws/docker/library/python:3.11-slim
 
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
 
WORKDIR /app
 
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
&& rm -rf /var/lib/apt/lists/*
 
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
 
COPY . .
 
# Expose port 5000 instead of 8501
EXPOSE 5000
 
# Run Streamlit on port 5000 without specifying server.address
CMD ["streamlit", "run", "app.py", "--server.port=5000"]

