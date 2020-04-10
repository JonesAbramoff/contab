VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Rotina_Importa_Dados(ByVal sNomeArqParam As String, ByVal objArqImpAux As ClassArqImportacaoAux) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Importa_Dados

    TelaAcompanhaBatch.Caption = "Acompanhamento da Importação"
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_IMPORTACAO_DADOS
      
    TelaAcompanhaBatch.Label1.Caption = "Número de Arquivos Importados:"
    TelaAcompanhaBatch.Label2.Caption = "Número de Registros Atualizado:"
    TelaAcompanhaBatch.Frame1.Caption = "Importação"
    TelaAcompanhaBatch.Frame2.Caption = "Atualização"
    
    Set TelaAcompanhaBatch.objImportacaoAux = objArqImpAux
            
    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatch.Show

    Rotina_Importa_Dados = SUCESSO

    Exit Function

Erro_Rotina_Importa_Dados:

    Rotina_Importa_Dados = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189750)

    End Select

    Exit Function

End Function

Public Function Rotina_Exporta_Dados(ByVal sNomeArqParam As String, ByVal objExportacaoAux As ClassArqExportacaoAux) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Exporta_Dados

    TelaAcompanhaBatch.Caption = "Acompanhamento da Exportação"
    TelaAcompanhaBatch.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatch.iRotinaBatch = ROTINA_EXPORTACAO_DADOS
        
    TelaAcompanhaBatch.Label1.Caption = "Número de Registros Exportados:"
    TelaAcompanhaBatch.Label2.Caption = "Número de Arquivos Gerados:"
    TelaAcompanhaBatch.Frame1.Caption = "Exportação"
    TelaAcompanhaBatch.Frame2.Caption = "Geração de Arquivos"
    
    Set TelaAcompanhaBatch.objExportacaoAux = objExportacaoAux

    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatch.Show

    Rotina_Exporta_Dados = SUCESSO

    Exit Function

Erro_Rotina_Exporta_Dados:

    Rotina_Exporta_Dados = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189751)

    End Select

    Exit Function

End Function
