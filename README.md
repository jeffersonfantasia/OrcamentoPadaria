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
```
Private Sub Workbook_Open()

    Planilha1.Select
    Range("b2").Value = ActiveWorkbook.Path
    Range("b3").Value = ActiveWorkbook.Name

End Sub
```


##### Criar novo módulo
```
Sub Atualizar()

    ' Torna a função volátil, ou seja, ela será recalculada sempre que houver qualquer alteração na planilha
    Application.Volatile
    
    ' Atualiza todas as conexões e consultas de dados
    ActiveWorkbook.RefreshAll
    
    ' Carimba a data e hora que finalizou a atualização
    Planilha1.Select
    Range("b4").Value = Now

    
    ' Mensagem final
    MsgBox "Atualização Concluida", vbInformation, "Status"


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