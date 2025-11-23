
# 📘 README — AssistenteQA

Assistente inteligente para análise de requisitos e geração automática de suítes de testes.

## 🚀 1. Pré-requisitos

### ✔ Python 3.10+
Baixe em https://www.python.org/downloads/

Durante a instalação, selecione:
✔ Add Python to PATH

## 🧱 2. Configurando o ambiente

### 📁 1. Baixar ou Extrair o arquivo AssistenteQA.zip
Baixe o zip e extraia tudo em uma pasta, exemplo:
```
C:\Projetos\MeuProjeto

### 🧪 2. Crie e ative o ambiente virtual

### Windows:
```
cd C:\Projetos\AssistenteQA
python -m venv venv (cria uma pasta com nome "Venv" que será usado como seu ambiente virtual)
venv\Scripts\activate (Comando para ativar o ambiente virtual)
```

### Linux/macOS:

cd AssistenteQA
python3 -m venv venv (cria uma pasta com nome "Venv" que será usado como seu ambiente virtual)
source venv/bin/activate (Comando para ativar o ambiente virtual)

## 📦 3. Instalar dependências

pip install -r requirements.txt

## 🔑 4. Configurar API KEY da Gemini

### Windows PowerShell:
```
setx API_KEY "SUA_CHAVE_AQUI"
```

### Linux/macOS:
```
export API_KEY="SUA_CHAVE_AQUI"
```

Ou crie `.env`:
```
API_KEY=SUA_CHAVE_AQUI
```

## ▶ 5. Executar o sistema

### FastAPI:
```
uvicorn main:app --reload
```

Acesse:
http://localhost:8000

### Interface local:
```
python app.py
```

---

## 🔍 6. Estrutura típica do projeto
```
AssistenteQA/
│── main.py
│── requirements.txt
│── README.md
│── utils/
│── templates/
│── static/
│── docs/
```

Pronto! Seu AssistenteQA está funcional.
