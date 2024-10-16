# Orçamento da Padaria do Seu Manoel
*Treinamento Black Eagle - Jovi Treinamentos!*

### Criando um orçamento empresarial automatizado utilizando Excel com Power Query e VBA

#### Início do Projeto

Comece criando os seguintes arquivos:
- **ParametrosGlobais.xlsx**: Trará os parâmetros para serem utilizado em qualquer um dos arquivos que serão criados. Esse arquivo será a nossa FUV (Fonte Única da Verdade)
- **TouroReprodutor.xlsm**: Será o nosso arquivo base para criar os demais arquivos `Trampolim.xlsm` em cada uma das pastas, já trazendo as macros, botões e vínculo com o arquivo `ParametrosGlobais.xlsx`

- Peça 1 - Faturamento
    - Definir a granularidade (Cliente/ Departamento / Marca / Família de produtos / Produto) por mês
    - Preço - no momento zero / Matriz de Reajustes
    - Prazo de recebimento - % recebido em cada período
    - Com isso entregaremos 2 tabelas: DRE e FLCX (planilha trampolim)
     Desabilite a conexão em segundo plano!


Abrir pasta em branco e iniciar a criação da estrutura de pastas:
- 001_FuncoesPadrao
- 010_ParametrosGlobais
    - Criar a tabela abaixo:
 - 020_Entidades
    - Plano de Contas
    - Centros de Custo
    - Produtos
    - Prazo de Recebimento
- 100_Faturamento
    - Projeção quantidade
    - Projeção de Preço
    - Prazos
    - Trampolim (arquivo que é atualizado)
- 150_TRIBUTOS
- 200_CMV
- ZZZ_Backup
    - TouroReprodutor - Macro - xlsm
      - Pasta Local: Range("b2").Value = ActiveWorkbook.Path
      - Arquivo Local: Range("b3").Value = ActiveWorkbook.Name
      - Data Atualização: Criar Módulo Atualizar() - Range("b4").Value = Now
  - Subir tabela para o power query como ParametrosLocais e criar os textos
        - vArquivoAtual
        - vPastaAtual
        - vPastaAnterior
   - Incluir os Parâmetros Globais e extratir as informações (deixar dinâmico o caminho
        - vDataInicio
        - vDataFim
     - Criar uma pasta para funções

- 999_Fim
    Planilha FIM que vai buscar o trampolim das demais pastas
    - Organizar os arquivos dentro das pastas BASES DRE e BASES FLCX no power quer
- Arquivo Atualizador.xlsm
    Onde estará a macro que vai rodar a atualização dos arquivos necessários.



