VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpCoinfoItemCRCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para contabilizacao automatica da importacao de faturas a receber do sistema coinfo

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjImpVoucher As ClassImpCoinfoItemFat
Private mvarobjFormCTBAux As Object

Const TRANSACAO_IMPORT_ITEMFAT_COINFO = "Vouchers"
Const MODELO_IMPORT_ITEMFAT_COINFO = "" '"Voucher CR Coinfo"
Const TRANSACAO_IMPORT_VOUCHER = 76

Private Const EMISSAO1 As String = "Data_Emissao"
Private Const VALOR1 As String = "Valor_Liquido"
Private Const NUMERO1 As String = "Numero_Voucher"
Private Const COMISSAO_RETIDA As String = "Comissao_Retida"
Private Const SERIE_VOU As String = "Serie_Voucher"
Private Const CLIENTE1 As String = "Codigo_Cliente"
Private Const CARTAO1 As String = "Pagto_Cartao"
Private Const UF_CLIENTE As String = "UF_Cliente"
Private Const CCL_VOU As String = "Ccl_Voucher"

Public Property Get iTransacao() As Integer
    iTransacao = TRANSACAO_IMPORT_VOUCHER
End Property

Public Property Get sNomeTransacao() As String
    sNomeTransacao = TRANSACAO_IMPORT_ITEMFAT_COINFO
End Property

Public Property Get sModelo() As String
    sModelo = MODELO_IMPORT_ITEMFAT_COINFO
End Property

Public Property Get Controls() As Object
Dim objControls As New Collection

    Set Controls = objControls

End Property

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

Dim lErro As Long, lDoc As Long, objPeriodo As New ClassPeriodo

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjImpVoucher = vParams(0)
    
    lErro = CF("Periodo_Le", gobjImpVoucher.dtData, objPeriodo)
    If lErro <> SUCESSO Then gError 130631
    
    lErro = CF("Voucher_Automatico_Trans", gobjImpVoucher.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, "CR", lDoc)
    If lErro <> SUCESSO Then gError 130630
    
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro2(Me, TRANSACAO_IMPORT_ITEMFAT_COINFO, gobjImpVoucher.lNumIntDoc, gobjImpVoucher.lCodEmp, FILIAL_MATRIZ, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjImpVoucher.iFilialEmpresa, 0, gobjImpVoucher.lNumVou, 0, "", 0, MODELO_IMPORT_ITEMFAT_COINFO)
    If lErro <> SUCESSO Then gError 32161
        
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case 32161, 130630, 130631
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184002)
     
    End Select
     
    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objEndereco As New ClassEndereco
Dim objFilialCliente As New ClassFilialCliente
Dim colCategoriaItem As New Collection
Dim objTipoCliente As New ClassTipoCliente
Dim colCategoriaItem1 As New Collection
Dim sItem As String
Dim objTipoClienteCategoria As ClassTipoClienteCategoria
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim iFilialCorporator As Integer
Dim iFilialCoinfo As Integer
Dim iAchou As Integer

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
            
        Case NUMERO1
            objMnemonicoValor.colValor.Add gobjImpVoucher.lNumVou
            
        Case EMISSAO1
            objMnemonicoValor.colValor.Add gobjImpVoucher.dtData
            
        Case VALOR1
            objMnemonicoValor.colValor.Add gobjImpVoucher.dValor
            
        Case COMISSAO_RETIDA
            objMnemonicoValor.colValor.Add gobjImpVoucher.dComissao
        
        Case SERIE_VOU
            objMnemonicoValor.colValor.Add gobjImpVoucher.sTipoDoc & " " & gobjImpVoucher.sSerie
        
        Case CLIENTE1
            objMnemonicoValor.colValor.Add gobjImpVoucher.lCodEmp
        
        Case CARTAO1
            objMnemonicoValor.colValor.Add gobjImpVoucher.iCartao
            
        Case UF_CLIENTE

            objCliente.lCodigo = gobjImpVoucher.lCodEmp
            
            'faz a leitura do cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 189432

            objEndereco.lCodigo = objCliente.lEndereco
        
            'Lê o endereço à partir do Código
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 189433
            
            objMnemonicoValor.colValor.Add objEndereco.sSiglaEstado
        
        Case CCL_VOU
            
            lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", gobjImpVoucher.lCodEmp, iFilialCorporator, iFilialCoinfo)
            If lErro <> SUCESSO Then gError 197258
            
            'se a filial for SP = 1 e codigo de grupo for 251 ==> filial = SP Interior = 7
            If iFilialCorporator = 1 And gobjImpVoucher.scodgru = "251" Then iFilialCorporator = 7
            
            objCliente.lCodigo = gobjImpVoucher.lCodEmp
            
            'faz a leitura do cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 197250
            
            If lErro <> SUCESSO Then gError 197251
            
            objFilialCliente.lCodCliente = gobjImpVoucher.lCodEmp
            objFilialCliente.iCodFilial = FILIAL_MATRIZ
            
            'Lê na tabela de FilialClienteCategorias, todas as categorias da Filial de Cliente (com o respectivo valor em cada categoria), guardando na coleção colCategoriaItem
            lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
            If lErro <> SUCESSO Then gError 197252
            
            iAchou = 0
            
            For Each objFilialCliCategoria In colCategoriaItem
            
                If objFilialCliCategoria.sCategoria = "CCL" Then
                    iAchou = 1
                    Exit For
                End If
            
            Next
        
            If iAchou = 0 Then gError 197253
        
            objTipoCliente.iCodigo = objCliente.iTipo
        
            'Lê na tabela de TiposDeClienteCategoria, todas as categorias de um Tipo de Cliente, guardando na coleção colCategoriaItem
            lErro = CF("TipoDeClienteCategorias_Le", objTipoCliente, colCategoriaItem1)
            If lErro <> SUCESSO Then gError 197254
        
            iAchou = 0
        
            For Each objTipoClienteCategoria In colCategoriaItem1
            
                If objTipoClienteCategoria.sCategoria = "CCL" Then
                    iAchou = 1
                    Exit For
                End If
            
            Next
        
            If iAchou = 0 Then gError 197255
        
            'retorna o item do produto/categoria
            lErro = CF("ProdutoCategoria_Le", gobjImpVoucher.sCodPro, "CCL", sItem)
            If lErro <> SUCESSO And lErro <> 92405 Then gError 197256
        
            If lErro <> SUCESSO Then gError 197257
        
            objMnemonicoValor.colValor.Add Format(iFilialCorporator, "@@") & Format(objTipoClienteCategoria.sItem, "@@") & Format(objFilialCliCategoria.sItem, "@@@") & Format(sItem, "@@")
        
        Case Else
            gError 39695
    
    End Select
    
    Calcula_Mnemonico = SUCESSO
     
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
     
    Select Case gErr
          
        Case 39695
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case 189432, 189433, 197250, 197252, 197254, 197256
            
        Case 197251
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            
        Case 197253
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLICATEGORIA_NAO_CADASTRADO", gErr, "CCL", objCliente.lCodigo)
            
        Case 197255
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTECATEGORIA_NAO_CADASTRADO", gErr, "CCL", objTipoCliente.iCodigo)

        Case 197257
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOCATEGORIA_NAO_CADASTRADO", gErr, "CCL", gobjImpVoucher.sCodPro)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184003)
     
    End Select
     
    Exit Function

End Function

Public Sub Encerrando()
    Set mvarobjFormCTBAux = Nothing
    Set gobjContabAutomatica = Nothing
    Set gobjImpVoucher = Nothing
End Sub

Private Sub Class_Terminate()
    Call Encerrando
End Sub


