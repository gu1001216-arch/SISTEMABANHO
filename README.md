# Sistema de Monitoramento — Tratamento de Peças (Pintura Eletrostática)

Controle em tempo real do fluxo **Preparação → Validação do líder → Banho**, com
histórico, gráficos, leitura de código de barras da OP e exportação para Excel.
Feito para rodar 24h na nuvem (Railway) com banco PostgreSQL.

---

## Como funciona o fluxo

1. **Operador de preparação** abre um card informando o número do cesto. O
   cronômetro de preparação começa na hora.
2. Ao terminar de encher o cesto, ele finaliza: escaneia o **código de barras da
   OP** (que autopreenche código, descrição e quantidade a partir da lista
   mestra — todos editáveis), escolhe o processo, o tipo (Normal/Retrabalho) e
   observações. O card vai para a fila do líder.
3. **O líder** vê o card, confere/ajusta o tempo de preparação e os dados, e
   libera para o banho.
4. **O operador de banho** vê o card na fila, inicia o banho (novo cronômetro) e
   finaliza quando o cesto sai. O card é concluído e entra no histórico.
5. **O admin** acompanha tudo no dashboard em tempo real: KPIs, gráficos, quadro
   de andamento e histórico completo, com download em Excel.

Todas as telas atualizam sozinhas a cada poucos segundos.

---

## Usuários e perfis

O sistema cria automaticamente estes usuários na primeira execução (troque as
senhas depois pela tela de Usuários):

| Login   | Senha     | Perfil  | Acessa                          |
|---------|-----------|---------|---------------------------------|
| `admin` | `admin123`| admin   | tudo (dashboard, usuários, mestra, Excel) |
| `lider` | `lider123`| lider   | tela do líder + preparação      |
| `banho` | `banho123`| banho   | tela do banho                   |
| `op1`…`op9` | `op1234` | prep | tela de preparação              |

Os 4 perfis:
- **prep** (operador de preparação): abre e finaliza cestos.
- **lider**: valida cards e libera ao banho (também pode abrir cestos).
- **banho** (operador de banho): controla a fila e o tempo de banho.
- **admin**: gerencia usuários, importa a lista mestra, vê dashboard e baixa Excel.

O admin gerencia tudo pela tela **Usuários**: adicionar, remover e trocar senha.
Os dados ficam no PostgreSQL, então **nunca se perdem** quando o app reinicia.

---

## Lista mestra (OP → código · descrição · quantidade)

O admin importa uma planilha **.xlsx** ou **.csv** na tela "Lista Mestra".
Colunas, nesta ordem: `OP`, `Código`, `Descrição`, `Quantidade`.
OPs já existentes são atualizadas; novas são adicionadas. Veja
`exemplo_lista_mestra.csv` no projeto.

Quando o operador escaneia o código de barras da OP, o sistema busca nessa lista
e preenche os campos automaticamente (sempre editáveis).

---

## Passo 1 — Subir para o GitHub

```bash
cd projeto
git init
git add .
git commit -m "Sistema de tratamento de peças"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/SEU_REPO.git
git push -u origin main
```

## Passo 2 — Publicar no Railway

1. Acesse railway.app e clique em **New Project → Deploy from GitHub repo**.
2. Selecione o repositório que você acabou de subir.
3. Ainda no projeto, clique em **New → Database → Add PostgreSQL**.
   O Railway cria a variável `DATABASE_URL` e o app a usa automaticamente.
4. Em **Variables** do serviço web, adicione uma variável:
   - `SECRET_KEY` = uma frase longa e aleatória (segurança das sessões).
5. O Railway detecta o `Procfile` e sobe com gunicorn. Quando o deploy terminar,
   clique em **Settings → Networking → Generate Domain** para ter a URL pública.

Pronto. Abra a URL, entre com `admin` / `admin123` e troque as senhas.

> **Importante:** sem o PostgreSQL conectado, o app cai para um SQLite local que
> **se perde** a cada reinício. Para uso 24/7, o passo 3 (adicionar PostgreSQL)
> é obrigatório.

---

## Rodar localmente (teste)

```bash
pip install -r requirements.txt
python app.py
# abre em http://localhost:5000
```

Sem `DATABASE_URL`, ele usa um arquivo SQLite local (`dados_local.db`) só para teste.

---

## Estrutura

```
projeto/
├── app.py                  # backend Flask + PostgreSQL + APIs
├── requirements.txt
├── Procfile                # comando de start (gunicorn) p/ Railway
├── railway.json            # config de build/deploy + restart automático
├── runtime.txt             # versão do Python
├── exemplo_lista_mestra.csv
├── static/
│   └── style.css
└── templates/
    ├── base.html  login.html  prep.html  lider.html
    ├── banho.html  dashboard.html  usuarios.html  mestre.html
```
