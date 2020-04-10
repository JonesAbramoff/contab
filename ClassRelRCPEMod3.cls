VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelRCPEMod3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para apoio a emissao de livro fiscal "Registro de Controle da Produção e do Estoque"

'Do livro Registro de Controle da
'Produção e do Estoque
'
'Art. 84. O livro Registro de Controle da Produção e do Estoque, modelo 3, Anexo I, destina-se à escrituração do documento fiscal e do documento de uso interno do estabelecimento, correspondentes à entrada e à saída, à produção e ao estoque de mercadorias.
'
'§ 1.º O livro Registro de Controle da Produção e do Estoque será utilizado pelo estabelecimento industrial ou a ele equiparado pela legislação federal e pelo estabelecimento atacadista, podendo, a critério do Fisco, ser exigido de estabelecimento de contribuinte de outro setor ou categoria, com as adaptações necessárias.
'
'§ 2.º Os lançamentos serão feitos operação a operação, devendo ser utilizada uma folha para cada espécie, marca, tipo e modelo de mercadoria.
'
'§ 3.º Os lançamentos serão feitos, nos quadros e colunas próprios, da seguinte forma:
'
'1. quadro "Produto": identificação da mercadoria, como definida no parágrafo anterior;
'
'2. quadro "Unidade": especificação da unidade (quilograma, metro, litro, dúzia etc.), de acordo com a legislação do IPI;
'
'3. quadro "Classificação Fiscal": indicação da posição, do inciso e do subinciso e alíquota, previstos na legislação do IPI, sendo que o disposto neste item não se aplica ao estabelecimento comercial não equiparado a industrial;
'
'4. colunas sob o título "Documento": espécie, série e subsérie, número de ordem e data da emissão do documento fiscal ou do documento de uso interno do estabelecimento, correspondente a cada operação;
'
'5. colunas sob o título "Lançamento": número e folha do livro Registro de Entradas ou do livro Registro de Saídas em que o documento fiscal tenha sido lançado, bem como a codificação fiscal, e, quando for o caso, a contábil;
'
'6. colunas sob o título "Entradas":
'
'a) coluna "Produção - No Próprio Estabelecimento": quantidade do produto industrializado no próprio estabelecimento;
'
'b) coluna "Produção - Em Outro Estabelecimento": quantidade do produto industrializado em outro estabelecimento da mesma empresa ou de terceiro, com mercadoria remetida para esse fim;
'
'c) coluna "Diversas": quantidade de mercadoria não classificada nas alíneas anteriores, inclusive a recebida de outro estabelecimento da mesma empresa ou de terceiro para industrialização e posterior retorno, consignando-se o fato, nesta última hipótese, na coluna "Observações";
'
'd) coluna "Valor": a base de cálculo do IPI, quando a entrada da mercadoria originar crédito desse tributo, ou se a entrada não gerar crédito ou quando se tratar de isenção, imunidade ou não incidência do mencionado tributo, será registrado o valor total atribuído à mercadoria;
'
'e) coluna "IPI": valor do imposto creditado, quando de direito;
'
'7. colunas sob o título "Saídas":
'
'a) coluna "Produção - No Próprio Estabelecimento": em se tratando de matéria-prima, produto intermediário e material de embalagem, a quantidade remetida do almoxarifado para o setor de fabricação, para industrialização no próprio estabelecimento, ou, em se tratando de produto acabado, a quantidade saída, a qualquer título, de produto industrializado no próprio estabelecimento;
'
'b) coluna "Produção - Em Outro Estabelecimento": em se tratando de matéria-prima, produto intermediário e material de embalagem, a quantidade saída para industrialização em outro estabelecimento da mesma empresa ou de terceiro, quando o produto industrializado deva retornar ao estabelecimento remetente, e, em se tratando de produto acabado, a quantidade saída, a qualquer título, de produto industrializado em estabelecimento de terceiro;
'
'c) coluna "Diversas": quantidade de mercadoria saída, a qualquer título, em hipótese não compreendida nas alíneas anteriores;
'
'd) coluna "Valor": base de cálculo do IPI, ou o valor total atribuído à mercadoria, em caso de saída com isenção ou não-incidência;
'
'e) coluna "IPI": valor do imposto, quando devido;
'
'8. coluna "Estoque": quantidade em estoque, após cada lançamento de entrada ou saída;
'
'9. coluna "Observações": anotações diversas.
'
'§ 4.º Não será escriturada neste livro a entrada de mercadoria destinada a integração ao ativo imobilizado ou para uso ou consumo do estabelecimento.
'
'§ 5.º Quando se tratar de produtos da mesma posição da tabela anexa ao Regulamento do IPI, pode o industrial ou o estabelecimento a ele equiparado agrupá-los numa mesma folha, desde que nesse sentido autorizados pela Secretaria da Receita Federal.
'
'§ 6.º O livro referido neste artigo pode, a critério da repartição fiscal de circunscrição, ser substituído por fichas, as quais devem ser:
'
'1. impressas com os mesmos elementos do livro substituído;
'
'2. numeradas tipograficamente, em ordem crescente de 1 a 999.999;
'
'3. impressas mediante AIDF, nos termos do artigo 7.º.
'
'§ 7.º Na hipótese do parágrafo anterior, deverá, ainda, ser previamente visada pela repartição fiscal de circunscrição a ficha-índice da utilização das fichas de controle da produção e do estoque, Anexo I, na qual, observada a ordem numérica crescente, será registrada a utilização de cada ficha, devendo ditas anotações estarem sempre atualizadas.
'
'§ 8.º A aprovação do Fisco, quanto à substituição do livro por fichas, considera-se formalizada com a concessão da autorização para sua impressão, na forma da Seção II, do Capítulo I, do Título III, no que couber.
'
'§ 9.º No último dia de cada mês, devem ser somadas as quantidades e valores constantes das colunas "Entradas" e "Saídas", acusando o saldo das quantidades em estoque, que será transportado para o mês seguinte.
'
'§ 10. A escrituração do livro de que trata este artigo ou das fichas referidas nos §§ 7.º e 8.º não pode atrasar por mais de 15 dias.
'
'Art. 85. É facultado ao contribuinte escriturar o livro, com as seguintes simplificações:
'
'I - lançamento de totais diários na coluna "Produção" - "No Próprio Estabelecimento"; sob o título "Entradas";
'
'II - lançamento de totais diários na coluna "Produção" - "No Próprio Estabelecimento", sob o título "Saídas", em se tratando de matéria-prima, produto intermediário e material de embalagem, quando remetidos do almoxarifado para o setor de fabricação, para industrialização no próprio estabelecimento;
'
'III - dispensa de escrituração das colunas sob os títulos "Documento" e "Lançamento", excetuada a coluna "Data" nos casos previstos nos incisos anteriores;
'
'IV - lançamento diário, em vez de após cada lançamento de entrada ou saída, na coluna "Estoque".
'
'§ 1.º O estabelecimento industrial ou a ele equiparado pela legislação do Imposto sobre Produtos Industrializados e o atacadista que possuir Controle Quantitativo de Mercadorias, capaz de permitir perfeita apuração do estoque permanente, pode utilizar, independentemente de autorização prévia, esse controle, em substituição ao livro modelo 3, desde que atendam às seguintes normas:
'
'1. comuniquem essa opção, por escrito, à repartição fiscal de circunscrição e, através do órgão da Secretaria da Receita Federal a que o estabelecimento optante estiver subordinado, à Superintendência Regional da Receita Federal, anexando modelos dos formulários adotados;
'
'2. apresentem à Fiscalização, quando solicitado, o Controle Quantitativo de Mercadorias substitutivo.
'
'§ 2.º A mercadoria que tenha pequena expressão na composição do produto final, tanto em termos físicos quanto em valor, poderá ser agrupada em uma só folha ou ficha, desde que se enquadre numa mesma posição da TIPI - Tabela de Incidência do Imposto sobre Produtos Industrializados.
'
'§ 3.º O estabelecimento atacadista não equiparado a produtor industrial, obrigado à adoção do livro modelo 3, fica dispensado da escrituração das colunas "Valor" e "IPI", mantidas as outras simplificações.


Function ObterCodigo(iCodigo As Integer, ByVal iQuebra As Integer, ByVal sEntradaSaida As String, ByVal iTipoMov As Integer, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal iNatureza As Integer, ByVal iProdNaFilial As Integer) As Long
'retorna em iCodigo se o movimento deve ser considerado como:
'1: no próprio estabelecimento
'2: em outro estabelecimento
'3: diversos

'iQuebra: 1: nosso conosco, 2: nosso em 3os, 3: de 3os conosco

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_ObterCodigo

    iCodigo = 0
    
    Select Case iQuebra
    
        Case 1 'nosso conosco
    
            Select Case sEntradaSaida
            
                'entradas
                Case "E"
                
                    'se producao entrada => coluna 1
                    If iTipoMov = MOV_EST_PRODUCAO Then iCodigo = 1
                    
                    'se entrada de material beneficiado => coluna 2
                    If iTipoMov = 139 Then iCodigo = 2 '??? depois incluir novo tipo
                    
                    If iCodigo = 0 Then iCodigo = 3
                
                'saidas
                Case Else
                
                    'se nao foi informado se o produto é produzido na filial
                    If iProdNaFilial = -1 Then
                                            
                        'Critica se o produto está cadastrado nessa Filial
                        objProdutoFilial.sProduto = sProduto
                        objProdutoFilial.iFilialEmpresa = iFilialEmpresa
                        
                        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                        If lErro <> SUCESSO And lErro <> 28261 Then gError 130491
                        
                        'Se Produto não está cadastrado na Filial, erro
                        If lErro = 28261 Then gError 130492
                                            
                    End If
                    
                    Select Case iNatureza
                    
                        'se produto acabado
                        Case NATUREZA_PROD_PRODUTO_ACABADO
                    
                            'se produzido na propria filial
                            If iProdNaFilial <> 0 Then
                                iCodigo = 1
                            Else
                                iCodigo = 2
                            End If
                            
                        Case NATUREZA_PROD_MATERIA_PRIMA, NATUREZA_PROD_PRODUTO_INTERMEDIARIO, NATUREZA_PROD_EMBALAGENS
                        
                            'se requisicao para producao => 1
                            If iTipoMov = MOV_EST_REQ_PRODUCAO Then iCodigo = 1
                            
                            'se remessa de material para beneficiamento => 2
                            If iTipoMov = MOV_EST_MAT_NOSSO_PARA_BENEF_SAIDA Then iCodigo = 2
                            
                            If iCodigo = 0 Then iCodigo = 3
        
                        Case Else
                        
                            iCodigo = 3
        
                    End Select
            
            End Select
        
        Case 2 'nosso em 3os
            
            iCodigo = 3
            
        Case 3 'de 3os conosco
        
            iCodigo = 3
    
    End Select
    
    ObterCodigo = SUCESSO
     
    Exit Function
    
Erro_ObterCodigo:

    ObterCodigo = gErr
     
    Select Case gErr
          
        Case 130491
        
        Case 130492
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_NAO_CADASTRADO", gErr, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130493)
     
    End Select
     
    Exit Function

End Function
