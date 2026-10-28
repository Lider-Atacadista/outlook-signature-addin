# 📧 Suplemento de Assinatura Corporativa (Office 365)
**Lider Atacadista (`@lideratacadista.com.br`) & Bem Família (`@bemfamilia.ind.br`)**

Compatível nativamente com o **Novo Outlook**, **Outlook na Web (OWA)** e **Outlook Clássico**.

---

## 📁 Estrutura dos Arquivos

* `manifest.xml`: Manifesto que você envia no portal de Administração do Microsoft 365.
* `autorun.html` e `autorun.js`: Código em segundo plano que escuta o evento de novo e-mail (`OnNewMessageCompose`) e injeta o HTML da assinatura.
* `index.html`: Página inicial com o status do serviço.
* `users.json`: Catálogo com os dados cadastrais (Cargo, Departamento, Telefones) de cada e-mail.
* `Atualizar-UsuariosJson.ps1`: Script PowerShell para exportar os dados atualizados do Active Directory diretamente para o `users.json`.

---

## 🚀 Passo a Passo de Implantação

### 1️⃣ Criar o Azure Static Web App (Gratuito)

1. Acesse o **Portal do Azure** ([portal.azure.com](https://portal.azure.com)).
2. Na barra de pesquisa superior, digite **Static Web Apps** e clique em **Criar (Create)**.
3. Preencha as opções básicas:
   * **Assinatura / Grupo de Recursos:** Selecione a sua assinatura existente.
   * **Nome:** ex: `assinatura-lider`
   * **Plano:** `Free (Gratuito)`
   * **Região:** `East US 2` ou `Brazil South`
   * **Origem da Implantação (Deployment source):**
     * Se usar GitHub: conecte ao seu repositório.
     * Se preferir subir direto sem Git: selecione **"Outro (Other)"** e use o Azure CLI / SWA CLI ou crie vazio.
4. Clique em **Revisar + Criar**.
5. Quando criar, copie a URL gerada pela Microsoft (ex: `https://gentle-wave-01234567.azurestaticapps.net`).

---

### 2️⃣ Publicar os Arquivos no Azure

Os arquivos que devem ser publicados na raiz do site são:
* `index.html`
* `autorun.html`
* `autorun.js`
* `users.json`

*(Você pode subir via GitHub Actions, via extensão do VS Code "Azure Static Web Apps" clicando com botão direito na pasta e escolhendo "Deploy to Static Web App", ou via comando `swa deploy`).*

---

### 3️⃣ Atualizar o `manifest.xml`

Abra o arquivo `manifest.xml` e localize a linha:
```xml
<bt:Url id="autorun.Url" DefaultValue="https://SUA-APP.azurestaticapps.net/autorun.html"/>
```
Substitua `https://SUA-APP.azurestaticapps.net` pela URL real gerada pelo Azure (ex: `https://gentle-wave-01234567.azurestaticapps.net/autorun.html`) e salve o arquivo.

---

### 4️⃣ Publicar no Microsoft 365 Admin Center

1. Acesse o **Centro de Administração do Microsoft 365** ([admin.microsoft.com](https://admin.microsoft.com)) com sua conta de Administrador Global.
2. No menu lateral esquerdo, vá em:
   **Configurações (Settings) > Aplicativos Integrados (Integrated Apps)**.
3. Clique na aba superior **Suplementos (Add-ins)**.
4. Clique em **Implantar Suplemento (Deploy Add-in)** e clique em **Avançar**.
5. Selecione a opção: **"Carregar aplicativos personalizados" (Upload custom apps)**.
6. Escolha **"Tenho o arquivo de manifesto neste dispositivo"** e selecione o seu `manifest.xml`.
7. Clique em **Carregar**.
8. Na tela de atribuição de usuários:
   * Para testar com a equipe de TI: selecione **"Usuários/Grupos específicos"** e adicione o seu usuário.
   * Para liberar para a empresa toda: selecione **"Toda a organização"**.
9. Método de implantação: escolha **"Fixo (Obrigatório)"**.
10. Conclua a implantação!

---

### 5️⃣ Como Manter os Cargos e Ramais Atualizados

O script `Atualizar-UsuariosJson.ps1` pode ser colocado para rodar 1x por dia (via Tarefa Agendada no servidor do Active Directory):
* Ele lê as alterações de funcionários, cargos e telefones no AD.
* Sobrescreve o `users.json`.
* Com isso, ninguém precisa configurar nada manualmente: ao abrir o Outlook, os dados sempre estarão atualizados!
