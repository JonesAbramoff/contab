VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaEscopo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

Public Property Get EscDescricao() As Object
     Set EscDescricao = objUserControl.Controls("EscDescricao")
End Property

Public Property Get EscExpectativa() As Object
     Set EscExpectativa = objUserControl.Controls("EscExpectativa")
End Property

Public Property Get EscFatores() As Object
     Set EscFatores = objUserControl.Controls("EscFatores")
End Property

Public Property Get EscRestricoes() As Object
     Set EscRestricoes = objUserControl.Controls("EscRestricoes")
End Property

Public Property Get EscPremissas() As Object
     Set EscPremissas = objUserControl.Controls("EscPremissas")
End Property

Public Property Get EscExclusoes() As Object
     Set EscExclusoes = objUserControl.Controls("EscExclusoes")
End Property

Public Function Move_Tela_Memoria(ByVal objTela As Object, ByVal objPRJEscopo As ClassPRJEscopo) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objPRJEscopo.sDescricao = EscDescricao.Text
    objPRJEscopo.sExpectativa = EscExpectativa.Text
    objPRJEscopo.sFatoresSucesso = EscFatores.Text
    objPRJEscopo.sRestricoes = EscRestricoes.Text
    objPRJEscopo.sPremissas = EscPremissas.Text
    objPRJEscopo.sExclusoesEspecificas = EscExclusoes.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181434)

    End Select

    Exit Function

End Function

Function Traz_PRJEscopo_Tela(ByVal objTela As Object, ByVal objPRJEscopo As ClassPRJEscopo) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_PRJEscopo_Tela

    'Lê o PRJEscopo que está sendo Passado
    lErro = CF("PRJEscopo_Le", objPRJEscopo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181435

    If lErro = SUCESSO Then

       EscDescricao.Text = objPRJEscopo.sDescricao
        EscExpectativa.Text = objPRJEscopo.sExpectativa
        EscFatores.Text = objPRJEscopo.sFatoresSucesso
        EscRestricoes.Text = objPRJEscopo.sRestricoes
        EscPremissas.Text = objPRJEscopo.sPremissas
        EscExclusoes.Text = objPRJEscopo.sExclusoesEspecificas

    End If

    Traz_PRJEscopo_Tela = SUCESSO

    Exit Function

Erro_Traz_PRJEscopo_Tela:

    Traz_PRJEscopo_Tela = gErr

    Select Case gErr

        Case 181435

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181436)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()
    Set m_objUserControl = Nothing
End Sub
