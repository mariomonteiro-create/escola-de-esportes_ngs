# SGE MasterPro Web 🏆
**Sistema de Gestão Esportiva — Colégio 7 de Setembro**
Versão Web (Flask + HTML/JS) — multi-usuário, tempo real

---

## ⚡ Como rodar

### 1. Instalar dependências
```bash
pip install flask flask-cors
```

### 2. Copiar o banco de dados antigo (opcional)
Se quiser manter os dados do sistema desktop, coloque o arquivo
`SGE_MasterPro_V12.db` na mesma pasta que o `app.py`.
Se não houver banco, ele será criado automaticamente.

### 3. Iniciar o servidor
```bash
python app.py
```
O sistema estará disponível em: **http://SEU-IP:5000**

---

## 🌐 Acesso em rede local
Para que **outras pessoas na rede** acessem, use o IP da máquina:
- Windows: `ipconfig` → encontre o IPv4 (ex: 192.168.1.10)
- Mac/Linux: `ifconfig` ou `ip addr`
- Acesse: `http://192.168.1.10:5000`

---

## 🔐 Usuários padrão
| Usuário | Senha    |
|---------|----------|
| admin   | c7s2026  |
| italo   | esporte  |

---

## 📦 Estrutura de arquivos
```
sge_web/
├── app.py              ← Backend Flask (servidor)
├── requirements.txt    ← Dependências Python
├── SGE_MasterPro_V12.db ← Banco de dados SQLite
└── templates/
    └── index.html      ← Frontend completo (SPA)
```

---

## ✅ Funcionalidades
- 👥 **Atletas & Técnicos** — Cadastro, edição, exclusão com filtros
- ⚽ **Jogos** — Agenda completa, registro de placar
- 📋 **Convocação** — Gestão de status por jogo (Convocado/Presente/Ausente/Liberado)
- 📊 **Estatísticas** — V/E/D, aproveitamento, gols
- 👕 **Estoque** — Uniformes com tamanhos, entradas e vendas
- 🏃 **Atividades** — Grade de escolinhas, seleções e ed. física
- 📨 **Comunicação** — Textos prontos para direção e transporte
- ⚙️ **Configurações** — Listas globais e assinaturas

---

## 🚀 Deploy em servidor (opcional)
Para uso 24/7 via internet, use **Render, Railway, Fly.io** ou um VPS.
Adicione `gunicorn` e um `Procfile`:
```
web: gunicorn app:app
```
