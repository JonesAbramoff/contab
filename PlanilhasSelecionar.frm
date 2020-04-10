VERSION 5.00
Begin VB.Form PlanilhasSelecionar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecionar Planilha"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "PlanilhasSelecionar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoOLAP 
      Caption         =   "Atualizar Cubos..."
      Height          =   1185
      Left            =   5685
      Picture         =   "PlanilhasSelecionar.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2670
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   510
      Left            =   5775
      Picture         =   "PlanilhasSelecionar.frx":0C98
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   345
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   510
      Left            =   5775
      Picture         =   "PlanilhasSelecionar.frx":0DF2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   975
      Width           =   855
   End
   Begin VB.ListBox ListPlanilhas 
      Height          =   2400
      ItemData        =   "PlanilhasSelecionar.frx":0EF4
      Left            =   135
      List            =   "PlanilhasSelecionar.frx":0EF6
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   345
      Width           =   5415
   End
   Begin VB.Label LabelDescricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   150
      TabIndex        =   4
      Top             =   2865
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Planilhas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   855
   End
End
Attribute VB_Name = "PlanilhasSelecionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjRelPlanilhas As AdmRelPlanilha

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Const PLANILHAS_CAMINHO = "C:\Planilhas\"

Private Sub BotaoCancela_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim sDBOlap As String

On Error GoTo Erro_Form_Load

    Set gobjRelPlanilhas = New AdmRelPlanilha
    
    BotaoOk.Enabled = False
    
    lErro = CF("Empresa_Le_DBOlap", sDBOlap)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192424
    
    If lErro = SUCESSO And Len(Trim(sDBOlap)) > 0 Then
    
        lErro = DBOLAP_Troca_ADM100(sDBOlap)
        If lErro <> SUCESSO Then gError 192425
        
    End If
         
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 192424, 192425
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164884)

    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(sSiglaModulo As String)
    
Dim lErro As Long
Dim colPlanilhas As New Collection
Dim sCodPlanilhas As String
Dim vntCodPlanilhas As Variant
Dim sGrupo As String
    
On Error GoTo Erro_Trata_Parametros

    lErro = Obter_Grupo(sGrupo)
    If lErro Then gError 12001
    
    'Preenche a colecao com os nomes dos relatorios existentes no BD
    lErro = CF("Planilhas_Le_GrupoModulo", colPlanilhas, sGrupo, sSiglaModulo)
    If lErro Then gError 12002
    
    For Each vntCodPlanilhas In colPlanilhas
        sCodPlanilhas = vntCodPlanilhas
        ListPlanilhas.AddItem (sCodPlanilhas)
    Next
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
            
        Case 12000
            Call Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_MODULO", gErr)
            
        Case 12001
            Call Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_GRUPO", gErr)
            
        Case 12002
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANILHAS2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164885)

    End Select
    
    Exit Function

End Function

Private Sub BotaoFechar_Click()
    Unload PlanilhasSelecionar
End Sub

Private Sub BotaoOK_Click()

Dim objApp As Object, sNomeXls As String
Dim sCaminho As String, sBuffer As String

On Error GoTo Erro_BotaoOK_Click
    
    sNomeXls = gobjRelPlanilhas.sNome
    
    'se o nome do xls nao contem o path completo
    If InStr(sNomeXls, "\") = 0 Then
    
        'buscar diretorio configurado
        sBuffer = String(128, 0)
        Call GetPrivateProfileString("Forprint", "DirXls", "c:\excel\", sBuffer, 128, "ADM100.INI")
        
        sBuffer = StringZ(sBuffer)
        If right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
        sCaminho = sBuffer & sNomeXls & ".xls"

    Else
        
        If UCase(right(sNomeXls, 4)) <> ".XLS" Then
            
            sCaminho = sNomeXls & ".xls"
            
        Else
        
            sCaminho = sNomeXls
            
        End If
    
    End If
    
    Set objApp = CreateObject("Excel.Application")
    objApp.Visible = True
    
    Call Shell(objApp.Path & "\Excel.exe " & sCaminho, 1)
        
    Unload PlanilhasSelecionar
    
    Exit Sub

Erro_BotaoOK_Click:
    
    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164886)

    End Select
           
    Exit Sub
   
End Sub

Private Sub ListPlanilhas_Click()
'pega o relatorio selecionado e mostrar a descricao

Dim objRelPlanilhas As New AdmRelPlanilha
Dim lErro As Long

On Error GoTo Erro_ListPlanilhas_Click
    
    If ListPlanilhas.ListIndex = -1 Then Exit Sub
    
    objRelPlanilhas.sCodPla = ListPlanilhas.Text
    lErro = CF("Planilhas_Le_CodPla", objRelPlanilhas)
       
    Select Case lErro
        
        Case AD_SQL_SUCESSO
            LabelDescricao.Caption = objRelPlanilhas.sDescricao
        
        Case AD_SQL_ERRO
            gError 12011
                
        Case AD_SQL_SEM_DADOS
            gError 12012
        
        Case Else
            gError 12013
        
        End Select
        
    Set gobjRelPlanilhas = objRelPlanilhas
        
    BotaoOk.Enabled = True
     
Exit Sub
    
Erro_ListPlanilhas_Click:

    lErro_Chama_Tela = gErr

    Select Case gErr
            
        Case 12011, 12012
             Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANILHAS2", gErr)
             
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164887)

    End Select
           
    Exit Sub
    
End Sub

Private Sub ListPlanilhas_DblClick()

Dim objRelPlanilhas As New AdmRelPlanilha
Dim lErro As Long

On Error GoTo Erro_ListPlanilhas_Dbl_Click
    
    If ListPlanilhas.ListIndex = -1 Then Exit Sub
    
    objRelPlanilhas.sCodPla = ListPlanilhas.Text
    lErro = CF("Planilhas_Le_CodPla", objRelPlanilhas)
    
    Select Case lErro
        
        Case AD_SQL_SUCESSO
            
            Set gobjRelPlanilhas = objRelPlanilhas
            Call BotaoOK_Click
              
        Case AD_SQL_ERRO
            gError 12014
            
        Case AD_SQL_SEM_DADOS
            gError 12015
            
        Case Else
            gError 12016
        
    End Select
     
    Exit Sub
    
Erro_ListPlanilhas_Dbl_Click:

    lErro_Chama_Tela = gErr

    Select Case gErr
            
        Case 12014, 12015
             Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANILHAS2", gErr)
             
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164888)

    End Select
           
    Exit Sub

End Sub

Private Sub LabelDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricao, Source, X, Y)
End Sub

Private Sub LabelDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub BotaoOLAP_Click()

On Error GoTo Erro_BotaoOLAP_Click

    If Hour(Now) < PROC_OLAP_HORARIO_INVALIDO_FIM And Hour(Now) >= PROC_OLAP_HORARIO_INVALIDO_INICIO Then gError 189563

    Call Chama_Tela("OLAP")
    
    Exit Sub
    
Erro_BotaoOLAP_Click:

    Select Case gErr
    
        Case 189563
            Call Rotina_Erro(vbOKOnly, "PROC_OLAP_HORARIO_INVALIDO", gErr, PROC_OLAP_HORARIO_INVALIDO_INICIO, PROC_OLAP_HORARIO_INVALIDO_FIM)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189564)

    End Select
    
    Exit Sub
    
End Sub

Private Function DBOLAP_Troca_ADM100(ByVal sBDOLAP As String)
    
Dim lErro As Long
Dim sOLAP As String
Dim sCnxStr As String
    
On Error GoTo Erro_DBOLAP_Troca_ADM100

    sOLAP = String(128, 0)
    Call GetPrivateProfileString("Forprint", "OLAP", "Demo", sOLAP, 128, "ADM100.INI")
    sOLAP = Replace(sOLAP, Chr(0), "")
    
    If UCase(sOLAP) <> UCase(sBDOLAP) Then
        Call WritePrivateProfileString("Forprint", "OLAP", sBDOLAP, "ADM100.INI")
    
        sCnxStr = String(1024, 0)
        Call GetPrivateProfileString("Forprint", "CnxStr", "", sCnxStr, 1024, "ADM100.INI")
        sCnxStr = Replace(sCnxStr, Chr(0), "")
        
        sCnxStr = Replace(sCnxStr, sOLAP, sBDOLAP)
        
        Call WritePrivateProfileString("Forprint", "CnxStr", sCnxStr, "ADM100.INI")
    
    End If
    
    DBOLAP_Troca_ADM100 = SUCESSO
    
    Exit Function
    
Erro_DBOLAP_Troca_ADM100:

    DBOLAP_Troca_ADM100 = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192426)

    End Select
    
    Exit Function

End Function
