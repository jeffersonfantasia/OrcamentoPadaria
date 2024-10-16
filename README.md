# Orçamento da Padaria do Seu Manoel
*Treinamento Black Eagle - Jovi Treinamentos!*

### Criando um orçamento empresarial automatizado utilizando Excel com Power Query e VBA

#### Início do Projeto

Comece criando os seguintes arquivos:
- **ParametrosGlobais.xlsx**: Trará os parâmetros para serem utilizado em qualquer um dos arquivos que serão criados. Esse arquivo será a nossa FUV (Fonte Única da Verdade)
- **TouroReprodutor.xlsm**: Será o nosso arquivo base para criar os demais arquivos `Trampolim.xlsm` em cada uma das pastas, já trazendo as macros, botões e vínculo com o arquivo `ParametrosGlobais.xlsx`

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

<br>

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

    
    'Mensagem final
    MsgBox "Atualização Concluida", vbInformation, "Status"


End Sub
```
<br>

> :warning: **Atenção:** Faça as alterações conforme necessidade

<br>




