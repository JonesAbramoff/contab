VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBxCRCtbAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para contabilizacao automatica das baixas de titulos a receber da AposFurnas

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gcolTitRec As Collection
Private gcolBxParcRec As ColBaixaParcRec
Private glNumIntBaixa As Long
Private gdtData As Date
Private gobjMovCCI As ClassMovContaCorrente

Private Const VALOR1 As String = "Valor_Baixar"

Private Const MODELO_CR_AF As String = "Importação"

Public Property Get Controls() As Object
Dim objControls As New Collection

    Set Controls = objControls

End Property

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

Dim lErro As Long, lDoc As Long, objPeriodo As New ClassPeriodo

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gcolBxParcRec = vParams(0)
    Set gcolTitRec = vParams(1)
    glNumIntBaixa = vParams(2)
    gdtData = vParams(3)
    Set gobjMovCCI = vParams(4)
    
    lErro = CF("Periodo_Le", gdtData, objPeriodo)
    If lErro <> SUCESSO Then gError 194077
    
    lErro = CF("Voucher_Automatico_Trans", giFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, "CR", lDoc)
    If lErro <> SUCESSO Then gError 194078
    
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro2(Me, "BaixaRec", glNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresa, 0, 0, 0, "", 1, MODELO_CR_AF)
    If lErro <> SUCESSO Then gError 194079
        
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case 194077 To 194079
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194080)
     
    End Select
     
    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, objFornecedor As New ClassFornecedor

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case VALOR1
            objMnemonicoValor.colValor.Add gobjMovCCI.dValor

        Case Else
            gError 194081
    
    End Select
    
    Calcula_Mnemonico = SUCESSO
     
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
     
    Select Case gErr
          
        Case 194081
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194082)
     
    End Select
     
    Exit Function

End Function

Public Sub Encerrando()
    Set gobjContabAutomatica = Nothing
    Set gcolTitRec = Nothing
    Set gcolBxParcRec = Nothing
    Set gobjMovCCI = Nothing
End Sub

Private Sub Class_Terminate()
    Call Encerrando
End Sub


