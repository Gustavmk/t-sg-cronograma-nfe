# Projeto automação de emails para cronograma de pagamentos

Através de uma listagem de emails em csv com delimitador ponto e vírgula (;), será necessário criar uma automação para envio de emails personalizados utilizando o template de email fornecido. A automação deve ser capaz de ler os dados do arquivo csv, substituir os placeholders no template com os valores correspondentes para cada destinatário e enviar os emails via Outlook Office365.

```csv
Nome,Email,ValorNF
João Silva;joão.silva@company.co;R$ 10.000,00
Maria Pereira;mariap@company.co;R$ 15.000,00
Carlos Souza;carlos.souza@company.co;R$ 20.000,00

```

## Definiçaõ do projeto

- Definir flow de como será feito e tecnologias: Quais agentes, quais tecnologias, quais passos, etc
- Scripts e trabalho
- Documentaçaõ e execução
- COntrole e revisão antes da execução
- monitoramento e ajustes pós execução
- Dê preferência ao Powershell ou Python 
- Sugira APIs para envio de email via Outlook Office365, como Microsoft Graph API, SendGrid ou SMTP. 

## Etapas do workflow

### 1. Ler o arquivo CSV
1. Implemente uma função para ler o arquivo CSV com delimitador ponto e vírgula (;) e extrair os dados necessários, como nome, email e valor da nota fiscal para cada destinatário.
2. Execute uma exceção caso tenha alguma falha na leitura do arquivo, como formato incorreto ou dados faltantes.

### 2. Substituir os placeholders no template de email com os valores correspondentes para cada destinatário

1. Certifique-se de substituir os placeholders no template com os valores correspondentes para cada destinatário, como o nome, valor da nota fiscal, mês atual, ano vigente, mês seguinte e ano do mês seguinte (se aplicável).

### 3. Enviar os emails personalizados via Outlook Office365

1. Execute uma automação para enviar emails personalizados para cada destinatário, utilizando os parâmetros fornecidos e o template de email acima. 
2. envio de email deverá ser feito via Outlook Office365.


### Template Email

Parametros: 
- MesAtual: mês atual, por extenso
- AnoVigente: ano vigente, por extenso
- MesSeguinte: mês seguinte ao mês atual, por extenso
- AnoMesSeguinte: ano do mês seguinte, caso seja diferente do ano vigente (após dezembro de um ano, o ano seguinte é diferente do ano vigente)
- ValorNF: tipo dinheiro formatdo, R$ XX.XXX,00
- 
```mardown
Boa tarde,

Segue cronograma referente ao mês de <MesAtual> de <AnoVigente>.

Valor da NF R$ <ValorNF>

Enviar NF e documentos para dp@twygo.com
- Cópia do pró-labore (MEI está isento).
- Comprovante de pagamento dos tributos (INSS pago) + Certidão Negativa de Débitos Federais

|Cronograma Ref. <MÊS>/<ANO>|
|--|
|Envio da Nota Fiscal até: 22/<MesAtual>/<AnoVigente> – 18:00 |
|Envio de documentos até: 22/<MesAtual>/<AnoVigente> – 18:00 |
|Data de pagamento: <quinto dia util>/<MesSeguinte>/<AnoMesSeguinte> |

 
Reforçamos a importância de que o envio da Nota Fiscal e dos demais documentos siga rigorosamente o cronograma estabelecido acima.
 
Conforme disposto na cláusula **2.7** do contrato, destacamos que:

> Ocorrendo atraso, pela CONTRATADA, na entrega da fatura e respectiva Nota Fiscal nos termos do item “h” do Preâmbulo, caso a emissão da nota fiscal ocorra a partir do dia 21 do mês da prestação de serviço, a data de pagamento será dia 20 do mês subsequente à prestação de serviço.

 
Portanto, o não cumprimento dos prazos acarretará o replanejamento da data de pagamento conforme previsto contratualmente.

Atenciosamente,
```

