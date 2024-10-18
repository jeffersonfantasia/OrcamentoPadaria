# Orçamento da Padaria do Seu Manoel
*Treinamento Black Eagle - Jovi Treinamentos!*

## Criando um orçamento empresarial automatizado utilizando Excel com Power Query e VBA

### Início do Projeto

Comece criando os seguintes arquivos:
- **ParametrosGlobais.xlsx**: Trará os parâmetros para serem utilizado em qualquer um dos arquivos que serão criados. Esse arquivo será a nossa FUV (Fonte Única da Verdade)
- **TouroReprodutor.xlsm**: Será o nosso arquivo base para criar os demais arquivos `ZZZ_Trampolim.xlsm` em cada uma das pastas, já trazendo as macros, botões e vínculo com o arquivo `ParametrosGlobais.xlsx`

<br>

| Parametros | Valor | Descricao
| ----------- | ----------- | ----------- | 
| Local | Automação VBA | Caminho onde o arquivo está salvo
| Arquivo | Automação VBA | Nome do arquivo
| Last Refresh | Automação VBA | Data e hora da última atualização

<br>

##### Workbook - Open
```vba
Private Sub Workbook_Open()

    Planilha1.Select
    Range("b2").Value = ActiveWorkbook.Path
    Range("b3").Value = ActiveWorkbook.Name

End Sub
```


##### Criar novo módulo
```vba
Sub Atualizar()

    Dim conexao As WorkbookConnection
    
    ' Torna a função volátil, ou seja, ela será recalculada sempre que houver qualquer alteração na planilha
    Application.Volatile
    

    ' Loop por todas as conexões no arquivo ativo
    For Each conexao In ActiveWorkbook.Connections
        ' Verifica se a conexão é uma conexão de consulta (tipo OLEDB ou ODBC)
        If conexao.Type = xlConnectionTypeODBC Or conexao.Type = xlConnectionTypeOLEDB Then
            ' Desabilita a atualização em segundo plano
            conexao.OLEDBConnection.BackgroundQuery = False
        End If
    Next conexao
    
    
    ' Atualiza todas as conexões e consultas de dados
    ActiveWorkbook.RefreshAll
    
    ' Carimba a data e hora que finalizou a atualização
    Planilha1.Select
    Range("b4").Value = Now

End Sub
```
<br>

> :warning: Faça as alterações conforme necessidade

<br>

Organizar as informações em pastas dentro do Power Query:
- Parametros Globais: Fazer referência a tabela trazendo a informação dos valores dos parâmetros no seu tipo primitivo (text, date, number...)
- Parametros Locais: Usar a informação do local do arquivo para automatizar a chamada do arquivo parametros globais.
- Funções: Incluir as funções necessárias da pasta 001_Funcoes
- Auditorias: Rodar a verificação de erro com a função `fxVerificaErros`
- Arquivos Base: Pasta vazia inicialmente para ser usado para alocar a ingestão dos arquivos excel nos arquivos `ZZZ_Trampolim.xlsm` que serão produzidos
- Fim: Para receber as etapas DRE e FLCX

<br>

> :warning: Desabilite a conexão em segundo plano das tabelas!

<br>

### Criando as entidades necessárias para o projeto
- Plano de Contas
- DRE
- Produto
- Receita (proporção entre os insumos para fazer 1 kg de pão, aprox. 20 pães)
- Filial

<br>

### Peça 1 - Faturamento
- Definição da granularidade: Produto
- Quantidade alvo por mês, ano, produto e filial
- Preço no momento zero / Matriz reajustes
- Prazo de recebimento: % recebido em cada período
- Finaliza com arquivo `ZZZ_Trampolim.xlsm` levando as informações de DRE e FLCX
    - Funções utilizadas: 
        - `fxListaUmaDataPorPeriodo` para gerar uma lista apenas do primeiro dia de cada mês entre a data inicial e final
        - `fxMultiplicacaoAcumulada` para gerar o reajuste acumulado para ser multiplicado pelo preço de venda
        - `fxQuebraPrazoPgto` para gerar uma lista com os dias de recebimento e seus devidos percentuais (não utilizada nesse projeto)
        - `fxTabelaFim` para gerar uma tabela com os dados esperados, mantendo assim a padronização do output
        - `fxExpandeTodasColunas` para expandir as colunas de uma tabela, sem precisar mencionar os nomes

<br>

### Peça 2 - Tributos
- Definição da alíquota já como parâmetro local do arquivo `ZZZ_Trampolim.xlsm`
- Puxa a informação de faturamento da peça anterior
- Finaliza com o arquivo `ZZZ_Trampolim.xlsm` levando as informações de DRE e FLCX

<br>

### Peça 3 - CMV
- Buscar a receita do produto acabado para saber a necessidade de compra
- Custos por produto, para isso devemos trazer a informação da quantidade de venda para estimar a compra
- Matriz de reajustes
- Prazo de pagamento ao fornecedor
- Finaliza com arquivo `ZZZ_Trampolim.xlsm` levando as informações de DRE e FLCX
    - Funções utilizadas:
        - `fxListaUmaDataPorPeriodo` para gerar uma lista apenas do primeiro dia de cada mês entre a data inicial e final
        - `fxQuebraPrazoPgto` para gerar uma lista com os dias de recebimento e seus devidos percentuais (não utilizada nesse projeto)
        - `fxMultiplicacaoAcumulada` para gerar o reajuste acumulado para ser multiplicado pelo preço de custo
        - `fxTabelaFim` para gerar uma tabela com os dados esperados, mantendo assim a padronização do output
        - `fxExpandeTodasColunas` para expandir as colunas de uma tabela, sem precisar mencionar os nomes

<br>

### Peça 4 - Despesas Variáveis
- Buscar o faturamento no arquivo `ZZZ_Trampolim.xlsm` da Peça 1
- Buscar o faturamento com a data de recebimento somente de cartão para cálculos das taxas
- Bascar a tabela de Plano de contas no arquivo `Tabelas.xlsx` para fazer a verificação se a conta informada realmente existe
- Finaliza com arquivo `ZZZ_Trampolim.xlsm` levando as informações de DRE e FLCX
    - Funções utilizadas:
        - `fxTabelaFim` para gerar uma tabela com os dados esperados, mantendo assim a padronização do output
        - `fxExpandeTodasColunas` para expandir as colunas de uma tabela, sem precisar mencionar os nomes

<br>

### Peça 5 - Despesas Fixas
- Realizar o lançamentos das despesas fixas separado por filial, conta, centro de custo
- Bascar a tabela de Plano de contas no arquivo `Tabelas.xlsx` para fazer a verificação se a conta informada realmente existe
- Finaliza com arquivo `ZZZ_Trampolim.xlsm` levando as informações de DRE e FLCX
    - Funções utilizadas:
        - `fxTabelaFim` para gerar uma tabela com os dados esperados, mantendo assim a padronização do output
        - `fxExpandeTodasColunas` para expandir as colunas de uma tabela, sem precisar mencionar os nomes

### FIM
- Nessa etapa iremos buscar todos os arquivos DRE e FLCX criados
- Adicionaremos uma coluna com a origem do arquivo para caso haja discrepância nos valores conseguirmos debugar melhor o problema
- Assim finalizaremos com um `Table.Combine()` de todas as tabelas de DRE e outra com as tabelas FLCX,fechando assim o arquivo `Fim.xlsx`

<br>

### AUDITOR
- Agora criaremos um arquivo para ser o nosso auditor de erros, varrendo todos os arquivos da pasta raiz em busca daquilo que identificamos como possíveis erros:
- Exemplos:
    - Arquivos com caminhos não dinamizados
- Função utilizada: `fxVerificaTodosScriptsTodosArquivos`

<br>

### FINALIZANDO O PROJETO
- Criaremos o arquivo `Atualizador.xlsm` responsável por automatizar a atualização em cascata dos arquivos `ZZZ_Trampolim.xlsm` já criados.
- Vamos adicionar os parâmetros abaixo na tabela de parâmentros locais para termos a informação do tempo de atualização:
    - Data e Hora inicial da atualização
    - Data e Hora final da atualização
    - Tempo de atualização: Data Hora final - Data hora inicial
- Criaremos uma tabela com as informações necessárias para o script entender o que deve ser analisado

| Pasta | Arquivo | Atualiza
| ----------- | ----------- | ----------- | 
| Nomes das pastas | Nome do arquivo | SIM ou NAO

<br>

- Criar um novo módulo no VBA com as macros abaixo:

```vba

Sub SaveClose()
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

Sub AtualizaTudo()

    ' Torna a função volátil, ou seja, ela será recalculada sempre que houver qualquer alteração na planilha
    Application.Volatile
    
    ' Loop por todas as conexões no arquivo ativo
    For Each conexao In ActiveWorkbook.Connections
        ' Verifica se a conexão é uma conexão de consulta (tipo OLEDB ou ODBC)
        If conexao.Type = xlConnectionTypeODBC Or conexao.Type = xlConnectionTypeOLEDB Then
            ' Desabilita a atualização em segundo plano
            conexao.OLEDBConnection.BackgroundQuery = False
        End If
    Next conexao

    ' Atualiza todas as conexões e consultas de dados
    ActiveWorkbook.RefreshAll
    
    ' Carimba a data e hora que finalizou a atualização
    Planilha1.Select
    Range("b4").Value = Now
    
    AtualizarBases
    Planilha1.Select
    Range("b5").Value = Now()
    
    'Mensagem final
    MsgBox "Atualização Concluida", vbInformation, "Status"


End Sub


Sub AtualizarBases()
    
    ' === AGRADECIMENTO ESPECIAL A ALESSANDRO TROVATO ===
    Application.DisplayAlerts = False

    
    'Define pasta
    Planilha1.Select
    Pasta = Range("B2").Value
    
    Linha = 2
    
    'Volta para planilha com endereços
    Planilha3.Select
    
    While Cells(Linha, 3) <> ""
    
        If Cells(Linha, 3) = "SIM" Then
    
            'Carrega variáveis que compõe o endereço do arquivo a atualizar
            SubPasta = Cells(Linha, 1).Value
            Arquivo = Cells(Linha, 2).Value
            EnderecoCompleto = Pasta & "/" & SubPasta & "/" & Arquivo
            
            Workbooks.Open (EnderecoCompleto) 'Abre planilha
            'Quando abre a planilha, se houverem métodos .open serão executados
            
            'Vamos executar uma macro na planilha, mas caso não encontrar,
            'e der erro, vai prosseguir...
            On Error Resume Next
            Application.Run Arquivo & "!Atualizar" 'Roda essa macro que está nela
            
            'Verifica se a planilha atualizada contém erros, se sim, "JOÃO KLEBER!"
            Sheets("Erros").Select
            If Range("C2").Value = "" Then
                SaveClose
            Else:
                MsgBox "Erro identificado nesta planilha", vbCritical, "Verificar Erro"
                End
            End If
            
    
        End If
        Linha = Linha + 1
        
    Wend

    ' Restaura as configurações originais do Excel
    Application.DisplayAlerts = True

End Sub

```
<br>

O código da macro `AtualizaTudo()` é responsável por atualizar todas as conexões de dados do arquivo Excel, além de abrir e atualizar outras planilhas conforme uma lista específica. Vamos detalhar cada parte:

##### **Sub SaveClose()**
```vba
Sub SaveClose()
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub
```
Esta subrotina salva e fecha o arquivo do Excel que está atualmente ativo. Ela é usada mais à frente no processo de atualização de outras planilhas.

---

##### **Sub AtualizaTudo()**
```vba
Sub AtualizaTudo()

    Application.Volatile
```
Essa linha torna a função volátil, o que significa que ela será recalculada sempre que houver qualquer mudança na planilha.

```vba
    For Each conexao In ActiveWorkbook.Connections
        If conexao.Type = xlConnectionTypeODBC Or conexao.Type = xlConnectionTypeOLEDB Then
            conexao.OLEDBConnection.BackgroundQuery = False
        End If
    Next conexao
```
Aqui, o código percorre todas as conexões de dados no arquivo ativo e desabilita a atualização em segundo plano para conexões do tipo ODBC ou OLEDB, garantindo que as atualizações sejam feitas de forma síncrona.

```vba
    ActiveWorkbook.RefreshAll
```
Esta linha atualiza todas as conexões e consultas de dados do arquivo.

```vba
    Planilha1.Select
    Range("b4").Value = Now
```
Aqui, o código seleciona a "Planilha1" e insere a data e hora atuais na célula B4, indicando o momento em que a atualização foi concluída.

```vba
    AtualizarBases
    Range("b4").Value = Now()
```
A subrotina `AtualizarBases` é chamada para atualizar outras planilhas listadas em outra aba (a lógica dessa subrotina é explicada mais abaixo). Após essa atualização, a data e hora são novamente registradas na célula B4.

```vba
    MsgBox "Atualização Concluida", vbInformation, "Status"
```
Por fim, uma mensagem é exibida ao usuário informando que a atualização foi concluída.

---

##### **Sub AtualizarBases()**
Essa subrotina é o coração da lógica de atualização das outras planilhas.

```vba
    Application.DisplayAlerts = False
```
Aqui, a macro desabilita os alertas do Excel para evitar interrupções ou mensagens enquanto as planilhas estão sendo abertas e atualizadas.

```vba
    Planilha1.Select
    Pasta = Range("B2").Value
```
A "Planilha1" é selecionada, e o caminho da pasta onde estão os arquivos a serem atualizados é armazenado na variável `Pasta`, que vem da célula B2.

```vba
    Linha = 2
    Planilha3.Select
```
A variável `Linha` é inicializada com o valor 2 (referente à linha da planilha onde começa a lista de arquivos) e a "Planilha3" (onde parece estar a lista de arquivos) é selecionada.

```vba
    While Cells(Linha, 3) <> ""
        If Cells(Linha, 3) = "SIM" Then
```
O código entra em um laço `While`, que continua enquanto a célula na coluna 3 (coluna C) da linha atual não estiver vazia. Ele verifica se o valor da célula é "SIM", indicando que aquele arquivo precisa ser atualizado.

```vba
            SubPasta = Cells(Linha, 1).Value
            Arquivo = Cells(Linha, 2).Value
            EnderecoCompleto = Pasta & "/" & SubPasta & "/" & Arquivo
            Workbooks.Open (EnderecoCompleto)
```
Se o valor for "SIM", a macro coleta o nome da subpasta (coluna A), o nome do arquivo (coluna B), monta o caminho completo do arquivo e o abre.

```vba
            On Error Resume Next
            Application.Run Arquivo & "!Atualizar"
```
Aqui, a macro tenta rodar uma macro chamada "Atualizar" dentro do arquivo recém-aberto. Se houver um erro (como a macro não existir no arquivo), ele será ignorado por causa do `On Error Resume Next`.

```vba
            Sheets("Erros").Select
            If Range("C2").Value = "" Then
                SaveClose
            Else
                MsgBox "Erro identificado nesta planilha", vbCritical, "Verificar Erro"
                End
            End If
```
Depois, a macro verifica a planilha "Erros" do arquivo aberto. Se a célula C2 estiver vazia, significa que não houve erros, e o arquivo é salvo e fechado usando a subrotina `SaveClose`. Se C2 contiver algum valor, a macro alerta o usuário sobre o erro e interrompe o processo.

```vba
        Linha = Linha + 1
    Wend
```
A variável `Linha` é incrementada para que a macro processe o próximo arquivo na lista.

```vba
    Application.DisplayAlerts = True
```
Por fim, os alertas do Excel são reativados.

---

##### **Resumo**
A macro `AtualizaTudo` realiza a atualização das conexões de dados do arquivo atual, carimba a data de atualização, e também abre e atualiza outras planilhas com base em uma lista na "Planilha3". Ela desabilita alertas e rodará uma macro "Atualizar" em cada arquivo que abrir. Caso encontre algum erro em uma planilha, ela notifica o usuário e interrompe o processo.
