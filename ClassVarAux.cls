VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassVarAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'??? será que o objetivo deve ser obter/gravar a propria var ou apenas uma copia ????
    'a principio estou usando uma referencia p/proprio objeto
    
'para encapsular funcoes de tratamento de objetos da classe ClassVariavelCalculo

Private colVarsAtuais As New Collection         'colecao de VARS
'???Private colVarsAlteradas As New Collection      'colecao de VARS alteradas a partir do momento em que se iniciou o "atualizaimpostos"

Function ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'pesquisa em colVarsAtuais, se achar a variavel retorna-a, senao, retorna Nothing
Dim lErro As Long
Dim objVar As ClassVariavelCalculo, iAchou As Integer
On Error GoTo Erro_ObterVar

    iAchou = 0
    
    'percorre colVariavelCalculo
    For Each objVar In colVarsAtuais
        'se encontrar a variavel solicitada
        If sIdentificador = objVar.sIdentificador Then
            Set objVariavelCalculo = objVar
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Set objVariavelCalculo = New ClassVariavelCalculo
        objVariavelCalculo.sIdentificador = sIdentificador
        objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
        objVariavelCalculo.vValor = 0
    End If
    
    ObterVar = SUCESSO

    Exit Function

Erro_ObterVar:

    ObterVar = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154189)

    End Select

    Exit Function

End Function

Function GuardarVar(objVariavelCalculo As ClassVariavelCalculo)
'??? deve colocar o proprio objeto na colecao ou apenas uma copia ?
Dim objVar As ClassVariavelCalculo, iAchou As Integer

    iAchou = 0

    'procurar var em colVarsAtuais
    For Each objVar In colVarsAtuais
        'se encontrar a variavel solicitada
        If objVariavelCalculo.sIdentificador = objVar.sIdentificador Then
            iAchou = 1
            '???se o valor ou o tipo de preenchimento mudaram
            '???    incluir na colecao de colVarsAlteradas, se ainda nao existir lá

            'atualizar o valor e tipo em colVarsAtuais
            Call objVar.Copiar(objVariavelCalculo)
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        'incluir em colVarsAtuais
        Call colVarsAtuais.Add(objVariavelCalculo)
        ''???incluir em colVarsAlteradas
    End If
    
End Function

''Function ObterVarsAlteradas() As Long
'''retorna as vars que foram alteradas, excluindo as de uso interno
''
''End Function
''
''Function DesmarcarAlterados() As Long
'''esvazia a colecao de alterados
''End Function

