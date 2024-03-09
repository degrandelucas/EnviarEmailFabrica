# Automatização de Respostas por E-mail para Pedidos incompletos enviados pela Fábrica/Fornecedor

Este script em Python visa automatizar o processo de resposta por e-mail para pedidos incompletos enviados pela fábrica e/ou fornecedor. Ele examina a caixa de entrada de e-mails em busca de mensagens relacionadas a notas fiscais específicas e envia respostas personalizadas para cada uma delas, solicitando uma nova NF complementar para um determinado pedido.

## Funcionalidades Principais

- **Busca de E-mails Relacionados:** O script examina a caixa de entrada em busca de e-mails relacionados a notas fiscais pendentes de entrega.

- **Envio de Respostas Personalizadas:** Para cada e-mail encontrado, o script envia uma resposta personalizada solicitando uma nova NF complementar para o pedido correspondente.

## Passo a Passo do Projeto

1. **Conexão com Servidor IMAP:** O script conecta-se ao servidor IMAP para acessar a caixa de entrada de e-mails.

2. **Busca por E-mails Relacionados:** Ele busca por e-mails na caixa de entrada com base nos números de nota fiscal fornecidos.

3. **Envio de Respostas por E-mail:** Para cada e-mail encontrado, o script envia uma resposta personalizada solicitando uma nova NF complementar para o pedido correspondente.

## Tecnologias Utilizadas

- **Python:** Utilizado para desenvolver a lógica de automação do envio de respostas por e-mail.

- **imaplib:** Módulo utilizado para acessar e manipular caixas de entrada de e-mails via protocolo IMAP.

- **smtplib:** Módulo utilizado para enviar e-mails através do protocolo SMTP.

- **pandas:** Biblioteca utilizada para carregar e manipular dados de planilhas.

## Autor

- **Nome:** Lucas Degrande
- **Contato:** lucasdegrande15@gmail.com

