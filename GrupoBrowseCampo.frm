VERSION 5.00
Begin VB.Form GrupoBrowseCampo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Campos permitidos a Grupo x Tela"
   ClientHeight    =   5010
   ClientLeft      =   135
   ClientTop       =   1020
   ClientWidth     =   8250
   Icon            =   "GrupoBrowseCampo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6360
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   12
      Top             =   75
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "GrupoBrowseCampo.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "GrupoBrowseCampo.frx":02A4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "GrupoBrowseCampo.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Telas 
      Height          =   2790
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1995
      Width           =   3705
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2670
   End
   Begin VB.ComboBox Grupo 
      Height          =   315
      Left            =   975
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   1245
   End
   Begin VB.ListBox Campos 
      Height          =   2760
      Left            =   4320
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   2010
      Width           =   3705
   End
   Begin VB.ComboBox Arquivo 
      Height          =   315
      Left            =   5190
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1290
      Width           =   2850
   End
   Begin VB.Label Tela 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   1275
      Width           =   3165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Telas de Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   1740
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tela:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Módulo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   165
      TabIndex        =   9
      Top             =   765
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   255
      TabIndex        =   8
      Top             =   270
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Campos Permitidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4335
      TabIndex        =   7
      Top             =   1755
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4335
      TabIndex        =   6
      Top             =   1335
      Width           =   735
   End
End
Attribute VB_Name = "GrupoBrowseCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Limpa_Tela_Local()

    'Limpa ListBox Campos e ComboBox Arquivo
    Campos.Clear
    Arquivo.Clear

End Sub

Private Sub Arquivo_Click()

Dim lErro As Long
Dim colCampo As New Collection
Dim vCampo As Variant
Dim iIndice As Integer

On Error GoTo Erro_Arquivo_Click
   
    'Lê os nomes de Campos associados ao Arquivo
    lErro = Campos_Le3(Arquivo.Text, colCampo)
    If lErro Then Error 6545

    'Limpa ListBox Campos
    Campos.Clear
    
    'Preenche a ListBox Campos com os campos lidos no BD
    For Each vCampo In colCampo
        Campos.AddItem vCampo
    Next
    
    'Lê no BD os campos permitidos para Grupo,Tela,Arquivo selecionados
    Set colCampo = New Collection
    lErro = GrupoBrowseCampo_Le2(Grupo.Text, Tela.Caption, Arquivo.Text, colCampo)
    If lErro Then Error 6546
    
    'Seleciona na ListBox Campos os campos lidos no BD
    For Each vCampo In colCampo
        For iIndice = 0 To Campos.ListCount - 1
            
            If Campos.List(iIndice) = vCampo Then
                Campos.Selected(iIndice) = True
                Exit For
            End If
        
        Next
    Next
    
    Exit Sub

Erro_Arquivo_Click:

    Select Case Err

        Case 6545, 6546  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161717)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim colCampo As New Collection

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se Arquivo foi informado
    If Len(Arquivo.Text) = 0 Then Error 6547

    'Armazena campos selecionados em colCampo
    For iIndice = 0 To Campos.ListCount - 1
    
        If Campos.Selected(iIndice) Then
            colCampo.Add Campos.List(iIndice)
        End If
    
    Next
    
    'Se não há campo selecionado erro
    If colCampo.Count = 0 Then Error 6548
    
    'Grava colCampo no banco de dados
    lErro = GrupoBrowseCampo_Grava(Grupo.Text, Tela.Caption, Arquivo.Text, colCampo)
    If lErro Then Error 6549

    'Limpa ListBox Campos e ComboBox Arquivo
    Call Limpa_Tela_Local
    
    'Desseleciona Tela
    Telas.ListIndex = -1
    Tela.Caption = ""

Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 6547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", Err)
            Telas.SetFocus
            
        Case 6548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPO_NAO_SELECIONADO", Err)
            Campos.SetFocus

        Case 6549  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161718)

     End Select

     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    'Limpa ListBox Campos e ComboBox Arquivo
    Call Limpa_Tela_Local
    
    'Desseleciona Tela
    Telas.ListIndex = -1
    Tela.Caption = ""

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colGrupo As New Collection
Dim vGrupo As Variant
Dim colModulo As New Collection
Dim vModulo As Variant
   
On Error GoTo Erro_GrupoBrowseCampo_Form_Load
    
    Me.HelpContextID = IDH_CAMPOS_PERM_GRUPO_TELA
    
    'Lê Grupos existentes no BD
    lErro = Grupos_Le(colGrupo)
    If lErro = 6365 Then Error 6550
    If lErro Then Error 6551

    'Preenche List da ComboBox Grupo
    For Each vGrupo In colGrupo
        Grupo.AddItem (vGrupo)
    Next

    'Seleciona primeiro Grupo na ComboBox
    Grupo.ListIndex = 0

    'Lê Módulos existentes no BD
    lErro = Modulos_Le(colModulo)
    If lErro Then Error 6552

    'Preenche List da ComboBox Modulo
    For Each vModulo In colModulo
        Modulo.AddItem (vModulo)
    Next

    'Seleciona o primeiro Módulo na ComboBox Modulo
    Modulo.ListIndex = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_GrupoBrowseCampo_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err
        
        Case 6550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOS_NAO_CADASTRADOS", Err)
       
        Case 6551, 6552  'Tratado na rotina chamada
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161719)

    End Select

    Exit Sub

End Sub

Private Sub Grupo_Click()

Dim colCampo As New Collection
Dim vCampo As Variant
Dim iIndice As Integer
Dim lErro As Long
    
On Error GoTo Erro_Grupo_Click

    If Len(Arquivo.Text) > 0 Then    'Arquivo selecionado
    
        'Lê no BD os campos permitidos para Grupo,Tela,Arquivo selecionados
        lErro = GrupoBrowseCampo_Le2(Grupo.Text, Tela.Caption, Arquivo.Text, colCampo)
        If lErro Then Error 6553
        
        'Desseleciona todos os campos na ListBox Campos
        For iIndice = 0 To Campos.ListCount - 1
            Campos.Selected(iIndice) = False
        Next
                
        'Seleciona na ListBox Campos os campos lidos no BD
        For Each vCampo In colCampo
            For iIndice = 0 To Campos.ListCount - 1
                
                If Campos.List(iIndice) = vCampo Then
                    Campos.Selected(iIndice) = True
                    Exit For
                End If
            
            Next
        Next
    
    End If
    
    Exit Sub

Erro_Grupo_Click:

    Select Case Err

        Case 6553  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161720)

    End Select

    Exit Sub

End Sub

Private Sub Modulo_Click()

Dim lErro As Long
Dim colTela As New Collection
Dim vTela As Variant

On Error GoTo Erro_Modulo_Click
    
    'Limpa controles Campos, Arquivo e Tela
    Call Limpa_Tela_Local
    Telas.Clear
    Tela.Caption = ""
    
    'Lê os nomes de telas de Browse na tabela BrowseArquivo
    lErro = BrowseArquivo_Le_Telas(Modulo.Text, colTela)
    If lErro Then Error 6554
    
    'Preenche a list da ComboBox Tela com os nomes lidos
    For Each vTela In colTela
        Telas.AddItem vTela
    Next

    Exit Sub

Erro_Modulo_Click:

    Select Case Err

        Case 6554  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161721)

    End Select

    Exit Sub

End Sub

Private Sub Telas_DblClick()

Dim lErro As Long
Dim colArquivo As New Collection
Dim vArquivo As Variant

On Error GoTo Erro_Telas_DblClick

    'Escreve Tela selecionada na label Tela
    Tela.Caption = Telas.List(Telas.ListIndex)
    
    'Limpa controles Campos, Arquivo
    Call Limpa_Tela_Local
    
    'Lê os nomes de Arquivos associados à Tela na tabela BrowseArquivo
    lErro = BrowseArquivo_Le_Arquivos(Tela.Caption, colArquivo)
    If lErro Then Error 6555
    
    'Preenche a list da ComboBox Arquivo com os nomes lidos no BD
    For Each vArquivo In colArquivo
        Arquivo.AddItem vArquivo
    Next
    
    'Seleciona o primeiro arquivo na ComboBox Arquivo
    Arquivo.ListIndex = 0

    Exit Sub

Erro_Telas_DblClick:

    Select Case Err

        Case 6555  'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161722)

    End Select

    Exit Sub

End Sub
