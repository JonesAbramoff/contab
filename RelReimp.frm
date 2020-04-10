VERSION 5.00
Begin VB.Form Reimpressao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reimpressão"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "RelReimp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   510
      Left            =   5310
      Picture         =   "RelReimp.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   510
      Left            =   5310
      Picture         =   "RelReimp.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.ComboBox ListRelatorios 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   4005
   End
   Begin VB.Label DataArquivo 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1500
      TabIndex        =   6
      Top             =   1980
      Width           =   3570
   End
   Begin VB.Label LabelDescricaoRel 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   600
      TabIndex        =   5
      Top             =   735
      Width           =   4470
   End
   Begin VB.Label Label3 
      Caption         =   "Última Versão:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   2025
      Width           =   1290
   End
   Begin VB.Label Label1 
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
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   855
   End
End
Attribute VB_Name = "Reimpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ao carregar a tela desabilitar o ok
'preencher a tela como em relselecionar, qdo o usuario selecionar algum relatorio:
    '1)exibir a descricao
    '2)Criar funcao que com o nome do arquivo de reimpressao que fica no dicdados na tabela de relatorios, obter a data do arquivo de reimpressao (*.rei) usando a funcao do VB FileDateTime
    'se nao houver arquivo de reimpressao desabilitar o ok, se houver, habilita-lo e colocar a data por extenso no label.

Dim gobjRelSel As AdmRelSel

Private Sub BotaoCancela_Click()

    Unload Me
    
End Sub

Private Sub BotaoOK_Click()
    
    gobjRelSel.sCodRel = ListRelatorios.Text
    gobjRelSel.iCancela = 0
    Unload Reimpressao

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Reimpressao_Form_Load
    
    BotaoOK.Enabled = False
         
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Reimpressao_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173731)

    End Select
    
    Exit Sub

End Sub

Private Sub ListRelatorios_Click()
'Pega o relatório selecionado e mostra a descrição

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_ListRelatorios
    
    If ListRelatorios.ListIndex = -1 Then Exit Sub
    
    objRelatorio.sCodRel = ListRelatorios.Text
    lErro = CF("Relatorio_Le_Reimpressao",objRelatorio)
       
    Select Case lErro
        
        Case AD_SQL_SUCESSO
              LabelDescricaoRel.Caption = objRelatorio.sDescricao
              If objRelatorio.sNomeArqReimp <> "" Then
                DataArquivo.Caption = FileDateTime(objRelatorio.sNomeArqReimp)
                BotaoOK.Enabled = True
              Else
                BotaoOK.Enabled = False
                Exit Sub
              End If
               
        Case AD_SQL_ERRO
            Error 43302
                
        Case AD_SQL_SEM_DADOS
            Error 43303
        
        Case Else
            Error 43304
        
        End Select
        
    BotaoOK.Enabled = True
     
Exit Sub
    
Erro_ListRelatorios:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case 43302
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
        Case 43303
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173732)
    
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
    If lErro Then Error 43309
    
    'Preenche a coleção com os nomes dos relatórios existentes no BD
    lErro = CF("Relatorios_Le_GrupoModulo",colRelatorio, sGrupo, sSiglaModulo)
    If lErro Then Error 43310
    
    For Each vntCodRel In colRelatorio
        sCodRel = vntCodRel
        ListRelatorios.AddItem (sCodRel)
    Next
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
            
'        Case 12000
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_MODULO", Err)
            
        Case 43309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_GRUPO", Err)
            
        Case 43310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173733)

    End Select
    
    Exit Function

End Function


Private Sub DataArquivo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataArquivo, Source, X, Y)
End Sub

Private Sub DataArquivo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataArquivo, Button, Shift, X, Y)
End Sub

Private Sub LabelDescricaoRel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricaoRel, Source, X, Y)
End Sub

Private Sub LabelDescricaoRel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricaoRel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

