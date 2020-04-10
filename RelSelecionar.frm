VERSION 5.00
Begin VB.Form RelSelecionar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecionar Relatório"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "RelSelecionar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   510
      Left            =   5775
      Picture         =   "RelSelecionar.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   345
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   510
      Left            =   5775
      Picture         =   "RelSelecionar.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   975
      Width           =   855
   End
   Begin VB.ListBox ListRelatorios 
      Height          =   2400
      ItemData        =   "RelSelecionar.frx":03A6
      Left            =   135
      List            =   "RelSelecionar.frx":03A8
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
      Caption         =   "Relatório:"
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
Attribute VB_Name = "RelSelecionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'??? colocar no padrao de programacao
'??? deveria nao exibir relatorios tipo "printscreen" e TALVEZ outros que sejam disparados via programacao

Dim gobjRelSel As AdmRelSel

Private Sub BotaoCancela_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_SelecionaRel_Form_Load
    
    BotaoOK.Enabled = False
         
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_SelecionaRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173734)

    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(objRelSel As AdmRelSel, sSiglaModulo As String)
    
Dim lErro As Long
Dim colRelatorio As New Collection
Dim sCodRel As String
Dim vntCodRel As Variant
Dim sGrupo As String
    
On Error GoTo Erro_Trata_Parametros

    Set gobjRelSel = objRelSel
    
    gobjRelSel.iCancela = 1

    lErro = Obter_Grupo(sGrupo)
    If lErro Then Error 12001
    
    'Preenche a colecao com os nomes dos relatorios existentes no BD
    lErro = CF("Relatorios_Le_Outros",colRelatorio, sGrupo, sSiglaModulo)
    If lErro Then Error 12002
    
    For Each vntCodRel In colRelatorio
        sCodRel = vntCodRel
        ListRelatorios.AddItem (sCodRel)
    Next
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
            
        Case 12000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_MODULO", Err)
            
        Case 12001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_GRUPO", Err)
            
        Case 12002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173735)

    End Select
    
    Exit Function

End Function

Private Sub BotaoFechar_Click()
    Unload RelSelecionar
End Sub

Private Sub BotaoOK_Click()
    gobjRelSel.sCodRel = ListRelatorios.Text
    gobjRelSel.iCancela = 0
    Unload RelSelecionar
End Sub

Private Sub ListRelatorios_Click()
'pega o relatorio selecionado e mostrar a descricao

Dim objRelatorio As New AdmRelatorio
Dim lErro As Long

On Error GoTo Erro_ListRelatorios
    
    If ListRelatorios.ListIndex = -1 Then Exit Sub
    
    objRelatorio.sCodRel = ListRelatorios.Text
    lErro = CF("Relatorio_Le",objRelatorio)
       
    Select Case lErro
        
        Case AD_SQL_SUCESSO
              LabelDescricao.Caption = objRelatorio.sDescricao
        
        Case AD_SQL_ERRO
            Error 12011
                
        Case AD_SQL_SEM_DADOS
            Error 12012
        
        Case Else
            Error 12013
        
        End Select
        
    BotaoOK.Enabled = True
     
Exit Sub
    
Erro_ListRelatorios:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case 12011
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
        Case 12012
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173736)

    End Select
           
    Exit Sub
    
End Sub

Private Sub ListRelatorios_DblClick()

Dim objRelatorio As New AdmRelatorio
Dim lErro As Long

    On Error GoTo Erro_ListRelatorios_Dbl_Click
    
    If ListRelatorios.ListIndex = -1 Then Exit Sub
    
    objRelatorio.sCodRel = ListRelatorios.Text
    lErro = CF("Relatorio_Le",objRelatorio)
    
    Select Case lErro
        
        Case AD_SQL_SUCESSO
              BotaoOK_Click
        Case AD_SQL_ERRO
            Error 12014
        Case AD_SQL_SEM_DADOS
            Error 12015
        Case Else
            Error 12016
        
        End Select
     
Exit Sub
    
Erro_ListRelatorios_Dbl_Click:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case 12014
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
        Case 12015
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173737)

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

