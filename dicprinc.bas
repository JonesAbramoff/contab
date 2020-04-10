Attribute VB_Name = "Princ"
Option Explicit

Sub Main()

Dim lSistema As Long, lErro As Long
Dim objFlag As New AdmGenerico
Dim Y As New ClassConstCust

On Error GoTo Erro_Main

    'para permitir acessar o dicionario de dados
    lSistema = Sistema_Abrir()
    If lSistema = 0 Then Error 41657

    'carrega a tela p/identificacao do usuario
    Load DicLogin

    lErro = DicLogin.Trata_Parametros(objFlag)
    If lErro <> SUCESSO Then Error 41658

    DicLogin.Show vbModal

    If objFlag.vVariavel = False Then Error 41659
    
    Call Y.Inicializa_Tamanhos_String
    
    'carrega a tela principal
    Principal.Show
    
    Exit Sub
    
Erro_Main:

    Select Case Err
    
        Case 41657, 41658, 41659
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 159064)

    End Select
    
    If lSistema <> 0 Then Call Sistema_Fechar
    
    Exit Sub

End Sub

