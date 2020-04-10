VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPCTBVouInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para contabilizacao automatica da importacao de faturas a receber do sistema coinfo

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjVoucherInfo As ClassTRPVoucherInfo
Private mvarobjFormCTBAux As Object

Const TRANSACAO_VOUCHER = 81

Private Const EMISSAO1 As String = "Data_Emissao"
Private Const VALOR1 As String = "Valor_Doc"
Private Const NUMERO1 As String = "Numero_Voucher"
Private Const SERIE_VOU As String = "Serie_Voucher"
Private Const TIPO_VOU As String = "Tipo_Voucher"
Private Const TIPO_DOC As String = "Tipo_Doc"
Private Const CLIENTE1 As String = "Codigo_Cliente"
Private Const CARTAO1 As String = "Pagto_Cartao"
Private Const HISTORICO As String = "Historico"

Public Property Get iTransacao() As Integer
    iTransacao = TRANSACAO_VOUCHER
End Property

Public Property Get Controls() As Object
Dim objControls As New Collection

    Set Controls = objControls

End Property

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

Dim lErro As Long, lDoc As Long, objPeriodo As New ClassPeriodo

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjVoucherInfo = vParams(0)
    
    lErro = CF("Periodo_Le", gobjVoucherInfo.dtData, objPeriodo)
    If lErro <> SUCESSO Then gError 130631
    
    lErro = CF("Voucher_Automatico_Trans", gobjVoucherInfo.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, "CR", lDoc)
    If lErro <> SUCESSO Then gError 130630
    
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro2(Me, TRANSACAO_VOUCHER, gobjVoucherInfo.lNumIntDoc, gobjVoucherInfo.lCliForn, FILIAL_MATRIZ, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjVoucherInfo.iFilialEmpresa, 0, gobjVoucherInfo.lNumVou)
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
Dim iAchou As Integer

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
            
        Case NUMERO1
            objMnemonicoValor.colValor.Add gobjVoucherInfo.lNumVou
            
        Case EMISSAO1
            objMnemonicoValor.colValor.Add gobjVoucherInfo.dtData
            
        Case VALOR1
            objMnemonicoValor.colValor.Add gobjVoucherInfo.dValor
            
        Case SERIE_VOU
            objMnemonicoValor.colValor.Add gobjVoucherInfo.sSerie
        
        Case TIPO_VOU
            objMnemonicoValor.colValor.Add gobjVoucherInfo.sTipVou
        
        Case TIPO_DOC
            objMnemonicoValor.colValor.Add gobjVoucherInfo.sTipoDoc
        
        Case HISTORICO
            objMnemonicoValor.colValor.Add gobjVoucherInfo.sHistorico
        
        Case CLIENTE1
            objMnemonicoValor.colValor.Add gobjVoucherInfo.lCliForn
        
        Case CARTAO1
            objMnemonicoValor.colValor.Add gobjVoucherInfo.iCartao
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOCATEGORIA_NAO_CADASTRADO", gErr, "CCL", gobjVoucherInfo.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184003)
     
    End Select
     
    Exit Function

End Function

Public Sub Encerrando()
    Set mvarobjFormCTBAux = Nothing
    Set gobjContabAutomatica = Nothing
    Set gobjVoucherInfo = Nothing
End Sub

Private Sub Class_Terminate()
    Call Encerrando
End Sub


