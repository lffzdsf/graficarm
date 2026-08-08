# Publicar a API da Central

Esta pasta contém a ponte entre o GitHub Pages e o Google Sheets/Drive.

1. Abra https://script.google.com/ com a conta proprietária da planilha.
2. Crie um novo projeto chamado `Graficarm API`.
3. Substitua o conteúdo de `Code.gs` pelo arquivo desta pasta.
4. Em **Configurações do projeto**, marque a opção para mostrar o arquivo de manifesto e substitua `appsscript.json` pelo arquivo desta pasta.
5. Clique em **Implantar → Nova implantação → Aplicativo da Web**.
6. Configure **Executar como: Eu** e **Quem pode acessar: Qualquer pessoa**.
7. Autorize o acesso à planilha e ao Drive e copie a URL terminada em `/exec`.
8. Cole essa URL em `config.js`, no campo `apiUrl`, e publique a alteração.

O Web App permite leitura, criação, atualização e movimentação de demandas. A exclusão pública fica desativada por segurança. Imagens JPG/PNG/WebP e PDFs de até 8 MB são armazenados em `Workflow/Arquivos da Central/<ID>`.
