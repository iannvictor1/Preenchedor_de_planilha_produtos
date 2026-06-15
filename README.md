# Sistema de produtos

Aplicacao para consultar produtos com fotos, selecionar itens e gerar uma planilha `.xlsx` com imagens e fichas individuais.

O frontend principal e o React, localizado em `frontend-react`. A API fica em FastAPI e usa a logica compartilhada de `core.py`.

## Estrutura

- `backend.py`: API FastAPI.
- `core.py`: leitura de produtos, indexacao de fotos e geracao do Excel.
- `frontend-react/`: frontend React com Vite.
- `app.py`: interface Streamlit antiga, mantida como legado.
- `Fotos Cod/`: pasta padrao de fotos.
- `*.xlsx`: planilha modelo usada para as fichas.

## Como preparar as fotos

O sistema procura o codigo do produto no inicio do nome do arquivo.

Exemplos validos:

- `3335.jpg`
- `3335 (2).jpg`
- `3335-produto.png`

## Como rodar

Instale as dependencias Python:

```bash
pip install -r requirements.txt
```

Suba a API:

```bash
uvicorn backend:app --reload
```

Em outro terminal, entre na pasta do frontend e instale as dependencias:

```bash
cd frontend-react
npm install
```

Rode o frontend:

```bash
npm run dev
```

Abra o endereco mostrado pelo Vite, normalmente:

```text
http://127.0.0.1:5173
```

## Login, perfis e usuarios

Antes de iniciar o backend, configure as senhas e a chave usada para assinar as sessoes:

```powershell
$env:SELLER_PASSWORD="senha-do-vendedor"
$env:SUPERVISOR_PASSWORD="senha-do-supervisor"
$env:ADMIN_PASSWORD="senha-do-administrador"
$env:SESSION_SECRET="uma-chave-longa-e-aleatoria"
```

No primeiro inicio, essas variaveis criam as contas `vendedor`, `supervisor` e `admin`. Depois disso, o administrador pode criar, editar, desativar e excluir contas pela tela `Usuarios`. As senhas sao armazenadas com hash no arquivo local `users.db`, que nao deve ser versionado.

O vendedor pode manter ou aumentar o preco original; o supervisor e o administrador podem aumentar ou reduzir. A validacao e feita pela API antes da geracao do Excel.

## Atualizacao por Git em producao

Este projeto possui scripts no mesmo padrao do `bonificacao_system` para atualizar a producao com `git pull`.

Fluxo recomendado:

1. No computador de desenvolvimento, faca commit e push:

```powershell
git add .
git commit -m "descreva a alteracao"
git push
```

2. No computador de producao, atualize dentro da pasta do projeto:

```powershell
.\scripts\atualizar_producao.ps1 -ProjectDir "C:\sistema_produtos"
```

Se o backend/frontend rodam por tarefa agendada, informe o nome da tarefa:

```powershell
.\scripts\atualizar_producao.ps1 -ProjectDir "C:\sistema_produtos" -BackendTask "SistemaProdutosBackend" -FrontendTask "SistemaProdutosFrontend"
```

Por padrao, o script tenta reiniciar a tarefa principal chamada `Sistema Produtos`. Para usar outro nome:

```powershell
.\scripts\atualizar_producao.ps1 -ProjectDir "C:\sistema_produtos" -ProjectTask "Nome da tarefa"
```

Se rodam como servico do Windows:

```powershell
.\scripts\atualizar_producao.ps1 -ProjectDir "C:\sistema_produtos" -BackendService "SistemaProdutosBackend" -FrontendService "SistemaProdutosFrontend"
```

Para disparar do computador de desenvolvimento para a producao, habilite PowerShell Remoting/WinRM na maquina de producao e rode:

```powershell
.\scripts\atualizar_remoto.ps1 -ComputerName "NOME-OU-IP-DA-PRODUCAO" -ProjectDir "C:\sistema_produtos" -BackendTask "SistemaProdutosBackend"
```

O script faz:

- `git pull`
- cria `.venv` se ainda nao existir
- instala/atualiza dependencias Python
- roda `npm install`
- roda `npm run build`
- reinicia servicos ou tarefas agendadas informadas
- reinicia a tarefa principal `Sistema Produtos` por padrao
- se `-ProjectTask ""` for informado, tenta detectar e reiniciar automaticamente tarefas agendadas que apontem para a pasta do projeto

As pastas e arquivos de dados locais, como `Fotos Cod/`, `Fichas-*`, planilhas e CSVs, ficam fora do Git pelo `.gitignore`.

## Configuracao do frontend

Por padrao o React usa a API em:

```text
http://127.0.0.1:8000
```

Para mudar esse endereco, crie `frontend-react/.env.local` com:

```bash
VITE_API_URL=http://127.0.0.1:8000
```

## Saida gerada

A planilha criada contem:

- Aba `Produtos`: lista resumida dos produtos selecionados com imagem.
- Uma aba individual para cada produto selecionado, usando o modelo da ficha cadastral.

Se o nome da foto nao comecar com o codigo do produto, a foto nao sera encontrada automaticamente. Nesse caso, renomeie as imagens ou ajuste `image_code_from_name` em `core.py`.
