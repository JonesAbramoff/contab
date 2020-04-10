VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBatchFAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Rotina_Importa_Dados_Coinfo(ByVal sNomeArqParam As String, ByVal objFiltro As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Importa_Dados_Coinfo
   
    TelaAcompanhaBatchFAT.Caption = "Acompanhamento da Importação\Atualização dos dados da CoInfo"
    TelaAcompanhaBatchFAT.sNomeArqParam = sNomeArqParam
    
    Set TelaAcompanhaBatchFAT.gobjFiltro = objFiltro
    
    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchFAT.Show

    Rotina_Importa_Dados_Coinfo = SUCESSO

    Exit Function

Erro_Rotina_Importa_Dados_Coinfo:

    Rotina_Importa_Dados_Coinfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144654)

    End Select

    Exit Function

End Function


Public Function Rotina_GerComiInt(ByVal sNomeArqParam As String, objTRVGerComiInt As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_GerComiInt

    TelaAcompanhaBatchFAT1.Caption = "Acompanhamento da Geração de Comissões Internas"
    TelaAcompanhaBatchFAT1.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatchFAT1.iRotinaBatch = ROTINA_TRVGERACOMIINT_BATCH

    'Parametros da rotina
    Set TelaAcompanhaBatchFAT1.gobjTRVGerComiInt = objTRVGerComiInt

    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchFAT1.Show

    Rotina_GerComiInt = SUCESSO

    Exit Function

Erro_Rotina_GerComiInt:

    Rotina_GerComiInt = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197314)

    End Select

    Exit Function

End Function

