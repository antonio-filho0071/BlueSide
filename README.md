## **Gerador Automático de Ingressos e Envio por E-mail**

### **Descrição**
Este script Google Apps Script (GAS) automatiza a geração de ingressos personalizados em formato PDF a partir de dados em uma planilha Google Sheets. Após a geração, os ingressos são enviados por e-mail para os respectivos participantes.

### **Funcionalidades**
* **Leitura de dados:** Extrai informações como nome, e-mail e número de inscrição.
* **Geração de ingressos:** Cria cópias personalizadas de um modelo.
* **Envio de e-mails:** Envia e-mails com os ingressos anexados.

### **Como Utilizar**
1. Crie uma cópia deste script: Salve este código em um novo script no Google Apps Script.
2. Configure as constantes:
    templateId: Substitua pelo ID do seu modelo de apresentação no Google Slides.
    folderId: Substitua pelo ID da pasta onde deseja salvar os PDFs.
    X: Ajuste os índices X para corresponderem às colunas corretas na sua planilha (nome, e-mail, número de inscrição).

3. Adapte o conteúdo do e-mail: Personalize o assunto e o corpo do e-mail conforme necessário.
4. Execute a função: Chame a função gerarEEnviarIngressos() para iniciar o processo.


### **Observações:**

1. Substitua os placeholders: Certifique-se de substituir ID-GOOGLE-SLIDES, ID-PASTA-GOOGLEDRIVE, {{NOME}}, {{NUMERO_INSCRICAO}} e outros placeholders pelos valores corretos.
2. Ajuste os índices: Verifique os índices X para garantir que estejam apontando para as colunas corretas na sua planilha.
3. Teste o script: Execute o script em um ambiente de teste antes de utilizá-lo em produção.
