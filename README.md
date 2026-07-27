# Conectar o site ao Google Sheets

1. Envie `Dashboard_acompanhamento em campo.xlsx` para o Google Drive e abra como Google Sheets.
2. Na planilha, acesse **Extensões > Apps Script**.
3. Substitua o conteúdo do arquivo `Code.gs` pelo código desta pasta.
4. Se desejar, defina um valor em `APP_TOKEN`.
5. Clique em **Implantar > Nova implantação > Aplicativo da Web**.
6. Selecione uma permissão compatível com as regras da sua organização.
7. Copie a URL terminada em `/exec`.
8. No site, abra **Configuração**, cole a URL e o token e clique em **Testar**.

O site continua salvando localmente quando não houver internet. O botão
**Sincronizar** envia os registros pendentes assim que a conexão estiver disponível.

Para criar os gatilhos de relatórios automáticos, execute uma vez a função
`criarGatilhosDeRelatorio` no editor do Apps Script e autorize o acesso solicitado.
