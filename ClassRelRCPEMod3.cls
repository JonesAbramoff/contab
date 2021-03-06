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

'para apoio a emissao de livro fiscal "Registro de Controle da Produ��o e do Estoque"

'Do livro Registro de Controle da
'Produ��o e do Estoque
'
'Art. 84. O livro Registro de Controle da Produ��o e do Estoque, modelo 3, Anexo I, destina-se � escritura��o do documento fiscal e do documento de uso interno do estabelecimento, correspondentes � entrada e � sa�da, � produ��o e ao estoque de mercadorias.
'
'� 1.� O livro Registro de Controle da Produ��o e do Estoque ser� utilizado pelo estabelecimento industrial ou a ele equiparado pela legisla��o federal e pelo estabelecimento atacadista, podendo, a crit�rio do Fisco, ser exigido de estabelecimento de contribuinte de outro setor ou categoria, com as adapta��es necess�rias.
'
'� 2.� Os lan�amentos ser�o feitos opera��o a opera��o, devendo ser utilizada uma folha para cada esp�cie, marca, tipo e modelo de mercadoria.
'
'� 3.� Os lan�amentos ser�o feitos, nos quadros e colunas pr�prios, da seguinte forma:
'
'1. quadro "Produto": identifica��o da mercadoria, como definida no par�grafo anterior;
'
'2. quadro "Unidade": especifica��o da unidade (quilograma, metro, litro, d�zia etc.), de acordo com a legisla��o do IPI;
'
'3. quadro "Classifica��o Fiscal": indica��o da posi��o, do inciso e do subinciso e al�quota, previstos na legisla��o do IPI, sendo que o disposto neste item n�o se aplica ao estabelecimento comercial n�o equiparado a industrial;
'
'4. colunas sob o t�tulo "Documento": esp�cie, s�rie e subs�rie, n�mero de ordem e data da emiss�o do documento fiscal ou do documento de uso interno do estabelecimento, correspondente a cada opera��o;
'
'5. colunas sob o t�tulo "Lan�amento": n�mero e folha do livro Registro de Entradas ou do livro Registro de Sa�das em que o documento fiscal tenha sido lan�ado, bem como a codifica��o fiscal, e, quando for o caso, a cont�bil;
'
'6. colunas sob o t�tulo "Entradas":
'
'a) coluna "Produ��o - No Pr�prio Estabelecimento": quantidade do produto industrializado no pr�prio estabelecimento;
'
'b) coluna "Produ��o - Em Outro Estabelecimento": quantidade do produto industrializado em outro estabelecimento da mesma empresa ou de terceiro, com mercadoria remetida para esse fim;
'
'c) coluna "Diversas": quantidade de mercadoria n�o classificada nas al�neas anteriores, inclusive a recebida de outro estabelecimento da mesma empresa ou de terceiro para industrializa��o e posterior retorno, consignando-se o fato, nesta �ltima hip�tese, na coluna "Observa��es";
'
'd) coluna "Valor": a base de c�lculo do IPI, quando a entrada da mercadoria originar cr�dito desse tributo, ou se a entrada n�o gerar cr�dito ou quando se tratar de isen��o, imunidade ou n�o incid�ncia do mencionado tributo, ser� registrado o valor total atribu�do � mercadoria;
'
'e) coluna "IPI": valor do imposto creditado, quando de direito;
'
'7. colunas sob o t�tulo "Sa�das":
'
'a) coluna "Produ��o - No Pr�prio Estabelecimento": em se tratando de mat�ria-prima, produto intermedi�rio e material de embalagem, a quantidade remetida do almoxarifado para o setor de fabrica��o, para industrializa��o no pr�prio estabelecimento, ou, em se tratando de produto acabado, a quantidade sa�da, a qualquer t�tulo, de produto industrializado no pr�prio estabelecimento;
'
'b) coluna "Produ��o - Em Outro Estabelecimento": em se tratando de mat�ria-prima, produto intermedi�rio e material de embalagem, a quantidade sa�da para industrializa��o em outro estabelecimento da mesma empresa ou de terceiro, quando o produto industrializado deva retornar ao estabelecimento remetente, e, em se tratando de produto acabado, a quantidade sa�da, a qualquer t�tulo, de produto industrializado em estabelecimento de terceiro;
'
'c) coluna "Diversas": quantidade de mercadoria sa�da, a qualquer t�tulo, em hip�tese n�o compreendida nas al�neas anteriores;
'
'd) coluna "Valor": base de c�lculo do IPI, ou o valor total atribu�do � mercadoria, em caso de sa�da com isen��o ou n�o-incid�ncia;
'
'e) coluna "IPI": valor do imposto, quando devido;
'
'8. coluna "Estoque": quantidade em estoque, ap�s cada lan�amento de entrada ou sa�da;
'
'9. coluna "Observa��es": anota��es diversas.
'
'� 4.� N�o ser� escriturada neste livro a entrada de mercadoria destinada a integra��o ao ativo imobilizado ou para uso ou consumo do estabelecimento.
'
'� 5.� Quando se tratar de produtos da mesma posi��o da tabela anexa ao Regulamento do IPI, pode o industrial ou o estabelecimento a ele equiparado agrup�-los numa mesma folha, desde que nesse sentido autorizados pela Secretaria da Receita Federal.
'
'� 6.� O livro referido neste artigo pode, a crit�rio da reparti��o fiscal de circunscri��o, ser substitu�do por fichas, as quais devem ser:
'
'1. impressas com os mesmos elementos do livro substitu�do;
'
'2. numeradas tipograficamente, em ordem crescente de 1 a 999.999;
'
'3. impressas mediante AIDF, nos termos do artigo 7.�.
'
'� 7.� Na hip�tese do par�grafo anterior, dever�, ainda, ser previamente visada pela reparti��o fiscal de circunscri��o a ficha-�ndice da utiliza��o das fichas de controle da produ��o e do estoque, Anexo I, na qual, observada a ordem num�rica crescente, ser� registrada a utiliza��o de cada ficha, devendo ditas anota��es estarem sempre atualizadas.
'
'� 8.� A aprova��o do Fisco, quanto � substitui��o do livro por fichas, considera-se formalizada com a concess�o da autoriza��o para sua impress�o, na forma da Se��o II, do Cap�tulo I, do T�tulo III, no que couber.
'
'� 9.� No �ltimo dia de cada m�s, devem ser somadas as quantidades e valores constantes das colunas "Entradas" e "Sa�das", acusando o saldo das quantidades em estoque, que ser� transportado para o m�s seguinte.
'
'� 10. A escritura��o do livro de que trata este artigo ou das fichas referidas nos �� 7.� e 8.� n�o pode atrasar por mais de 15 dias.
'
'Art. 85. � facultado ao contribuinte escriturar o livro, com as seguintes simplifica��es:
'
'I - lan�amento de totais di�rios na coluna "Produ��o" - "No Pr�prio Estabelecimento"; sob o t�tulo "Entradas";
'
'II - lan�amento de totais di�rios na coluna "Produ��o" - "No Pr�prio Estabelecimento", sob o t�tulo "Sa�das", em se tratando de mat�ria-prima, produto intermedi�rio e material de embalagem, quando remetidos do almoxarifado para o setor de fabrica��o, para industrializa��o no pr�prio estabelecimento;
'
'III - dispensa de escritura��o das colunas sob os t�tulos "Documento" e "Lan�amento", excetuada a coluna "Data" nos casos previstos nos incisos anteriores;
'
'IV - lan�amento di�rio, em vez de ap�s cada lan�amento de entrada ou sa�da, na coluna "Estoque".
'
'� 1.� O estabelecimento industrial ou a ele equiparado pela legisla��o do Imposto sobre Produtos Industrializados e o atacadista que possuir Controle Quantitativo de Mercadorias, capaz de permitir perfeita apura��o do estoque permanente, pode utilizar, independentemente de autoriza��o pr�via, esse controle, em substitui��o ao livro modelo 3, desde que atendam �s seguintes normas:
'
'1. comuniquem essa op��o, por escrito, � reparti��o fiscal de circunscri��o e, atrav�s do �rg�o da Secretaria da Receita Federal a que o estabelecimento optante estiver subordinado, � Superintend�ncia Regional da Receita Federal, anexando modelos dos formul�rios adotados;
'
'2. apresentem � Fiscaliza��o, quando solicitado, o Controle Quantitativo de Mercadorias substitutivo.
'
'� 2.� A mercadoria que tenha pequena express�o na composi��o do produto final, tanto em termos f�sicos quanto em valor, poder� ser agrupada em uma s� folha ou ficha, desde que se enquadre numa mesma posi��o da TIPI - Tabela de Incid�ncia do Imposto sobre Produtos Industrializados.
'
'� 3.� O estabelecimento atacadista n�o equiparado a produtor industrial, obrigado � ado��o do livro modelo 3, fica dispensado da escritura��o das colunas "Valor" e "IPI", mantidas as outras simplifica��es.


Function ObterCodigo(iCodigo As Integer, ByVal iQuebra As Integer, ByVal sEntradaSaida As String, ByVal iTipoMov As Integer, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal iNatureza As Integer, ByVal iProdNaFilial As Integer) As Long
'retorna em iCodigo se o movimento deve ser considerado como:
'1: no pr�prio estabelecimento
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
                
                    'se nao foi informado se o produto � produzido na filial
                    If iProdNaFilial = -1 Then
                                            
                        'Critica se o produto est� cadastrado nessa Filial
                        objProdutoFilial.sProduto = sProduto
                        objProdutoFilial.iFilialEmpresa = iFilialEmpresa
                        
                        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                        If lErro <> SUCESSO And lErro <> 28261 Then gError 130491
                        
                        'Se Produto n�o est� cadastrado na Filial, erro
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
