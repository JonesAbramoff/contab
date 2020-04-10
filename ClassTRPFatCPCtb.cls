VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPCTBFatCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para contabilizacao automatica da importacao de comissoes over do sistema coinfo

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTitPag As ClassTituloPagar
Private mvarobjFormCTBAux As Object

Public sModelo As String

Private Const EMISSAO1 As String = "Data_Emissao"
Private Const VALOR1 As String = "Valor_da_Nota"
Private Const NUMERO1 As String = "Numero_Nota_Fiscal"

Public Property Get Controls() As Object
Dim objControls As New Collection
    Set Controls = objControls
End Property

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

Dim lErro As Long, lDoc As Long, objPeriodo As New ClassPeriodo

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjTitPag = vParams(0)
    
    lErro = CF("Periodo_Le", gobjTitPag.dtDataEmissao, objPeriodo)
    If lErro <> SUCESSO Then gError 130631
    
    lErro = CF("Voucher_Automatico_Trans", gobjTitPag.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, "CP", lDoc)
    If lErro <> SUCESSO Then gError 130630
    
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro2(Me, "NFFATPAG", gobjTitPag.lNumIntDoc, gobjTitPag.lFornecedor, gobjTitPag.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTitPag.iFilialEmpresa, 0, gobjTitPag.lNumTitulo, 0, "", 0, sModelo)
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

Dim lErro As Long, objFornecedor As New ClassFornecedor

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
            
        Case NUMERO1
            objMnemonicoValor.colValor.Add gobjTitPag.lNumTitulo
            
        Case EMISSAO1
            objMnemonicoValor.colValor.Add gobjTitPag.dtDataEmissao
            
        Case VALOR1
            objMnemonicoValor.colValor.Add gobjTitPag.dValorTotal

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
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184003)
     
    End Select
     
    Exit Function

End Function

Public Sub Encerrando()
    Set gobjContabAutomatica = Nothing
    Set gobjTitPag = Nothing
    Set mvarobjFormCTBAux = Nothing
End Sub

Private Sub Class_Initialize()
    sModelo = "FSCP"
End Sub

Private Sub Class_Terminate()
    Call Encerrando
End Sub



