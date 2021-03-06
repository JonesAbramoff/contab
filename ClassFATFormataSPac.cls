VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATFormataSPac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Inicializa_Tab_DadosFinanceiros_Usu(ByVal objTela As Object) As Long

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Inicializa_Tab_DadosFinanceiros_Usu
    
    'Guarda o c�digo do usu�rio no obj
    objUsuarios.sCodUsuario = gsUsuario
    
    'L� os dados desse usu�rio
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 127097
    
    'Se n�o encontrou o usu�rio => erro
    If lErro = 40832 Then gError 127098
    
    'Se o grupo do usu�rio � diferente de "FINANCEIRO" e "DIRETORIA" => bloqueia o acesso ao campo limite de cr�dito
    If UCase(objUsuarios.sCodGrupo) <> "DIRETORIA" And UCase(objUsuarios.sCodGrupo) <> "FINANCEIRO" Then objTela.Controls("LimiteCredito").Enabled = False
    
    Inicializa_Tab_DadosFinanceiros_Usu = SUCESSO
    
    Exit Function

Erro_Inicializa_Tab_DadosFinanceiros_Usu:

    Select Case gErr
    
        Case 127097
        
        Case 127098
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149743)
            
    End Select
    
End Function
