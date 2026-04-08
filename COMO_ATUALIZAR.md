# Como atualizar o dashboard diariamente

## Link do dashboard
Após configurar, o link será:
`https://SEU_USUARIO.github.io/pace-dashboard`

---

## Configuração inicial (só uma vez — 10 minutos)

### 1. Criar conta no GitHub
Acesse https://github.com e crie uma conta gratuita.

### 2. Criar repositório
- Clique em **New repository**
- Nome: `pace-dashboard`
- Marque **Public**
- Clique **Create repository**

### 3. Subir os arquivos
- Na página do repositório, clique **uploading an existing file**
- Arraste TODOS os arquivos desta pasta (incluindo as subpastas `.github/` e `relatorios/`)
- Clique **Commit changes**

### 4. Ativar GitHub Pages
- Vá em **Settings → Pages**
- Em "Source", selecione **GitHub Actions**
- Adicione este workflow: vá em **Actions → New workflow → set up a workflow yourself**
- Cole o conteúdo do arquivo `.github/workflows/update-dashboard.yml`
- Clique **Commit**

Após alguns segundos, o dashboard estará em:
`https://SEU_USUARIO.github.io/pace-dashboard`

---

## Atualização diária (2 minutos por dia)

1. Exporte os relatórios do Pipedrive normalmente (os mesmos 7 arquivos)
2. Acesse o repositório no GitHub
3. Clique na pasta **relatorios**
4. Clique **Add file → Upload files**
5. Arraste os novos arquivos e clique **Commit changes**
6. Em ~30 segundos o dashboard é atualizado automaticamente

**Não é necessário Python, Terminal ou nenhuma instalação.**

---

## Alternativamente: enviar para o Claude

Se preferir, basta enviar os arquivos ao Claude (como sempre fez).
O Claude gera o `index.html` atualizado e você faz upload via:
- GitHub (procedimento acima)
- Ou arraste para https://app.netlify.com/drop (gera link imediato)
