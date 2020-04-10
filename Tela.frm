VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Tela 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tela"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   1005
   ClientWidth     =   9330
   Icon            =   "Tela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin VB.Frame SSFrame1 
      Caption         =   "Tela"
      Height          =   3735
      Left            =   270
      TabIndex        =   11
      Top             =   630
      Width           =   5505
      Begin VB.Frame SSFrame2 
         Caption         =   "Customizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   405
         TabIndex        =   21
         Top             =   2385
         Width           =   4515
         Begin MSMask.MaskEdBox ClasseCust 
            Height          =   315
            Left            =   1455
            TabIndex        =   22
            Top             =   750
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProjetoCust 
            Height          =   315
            Left            =   1455
            TabIndex        =   23
            Top             =   300
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
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
            Left            =   570
            TabIndex        =   25
            Top             =   795
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   510
            TabIndex        =   24
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Original"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   390
         TabIndex        =   16
         Top             =   1155
         Width           =   4545
         Begin VB.Label ProjetoOrig 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1515
            TabIndex        =   20
            Top             =   270
            Width           =   2730
         End
         Begin VB.Label ClasseOrig 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1515
            TabIndex        =   19
            Top             =   690
            Width           =   2730
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   600
            TabIndex        =   18
            Top             =   330
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
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
            Left            =   630
            TabIndex        =   17
            Top             =   735
            Width           =   645
         End
      End
      Begin VB.TextBox Descricao 
         Height          =   315
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   750
         Width           =   2730
      End
      Begin VB.TextBox Nome 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   330
         Width           =   2730
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1125
         TabIndex        =   15
         Top             =   390
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   765
         TabIndex        =   14
         Top             =   765
         Width           =   945
      End
   End
   Begin VB.Frame SSFrame3 
      Caption         =   "Observação"
      Height          =   855
      Left            =   300
      TabIndex        =   9
      Top             =   4440
      Width           =   5475
      Begin VB.Label Label8 
         Caption         =   "A customização desta tela é a nível de Empresa. Para customizar a nível de Grupo clique o Botão  Tela x Grupo."
         Height          =   390
         Left            =   660
         TabIndex        =   10
         Top             =   300
         Width           =   4110
      End
   End
   Begin VB.CommandButton TelaxGrupo 
      Caption         =   "Tela x Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6780
      TabIndex        =   8
      Top             =   4650
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7440
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Tela.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "Tela.frx":02A4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "Tela.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Telas 
      Height          =   2985
      ItemData        =   "Tela.frx":0954
      Left            =   6060
      List            =   "Tela.frx":095B
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1380
      Width           =   3075
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   2250
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   195
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Telas"
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
      Left            =   6060
      TabIndex        =   3
      Top             =   1095
      Width           =   495
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
      Left            =   1365
      TabIndex        =   2
      Top             =   225
      Width           =   705
   End
End
Attribute VB_Name = "Tela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Limpa_Tela_Local()

    ProjetoOrig.Caption = ""
    ClasseOrig.Caption = ""

End Sub

Private Sub BotaoFechar_Click()

    Unload Tela

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objTela As New ClassDicTela

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se dados da Tela foram informados
    If Len(Nome.Text) = 0 Then Error 6355
    
    objTela.sNome = Nome.Text
    objTela.sDescricao = Trim(Descricao.Text)
    objTela.sProjeto_Customizado = Trim(ProjetoCust.Text)
    objTela.sClasse_Customizada = Trim(ClasseCust.Text)
        
    'grava a Tela no banco de dados (é um update)
    lErro = Tela_Grava(objTela)
    If lErro Then Error 6356
         
    'Limpa a Tela
    Call Limpa_Tela(Tela)
    Call Limpa_Tela_Local
  
Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 6355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_TELA_NAO_INFORMADO", Err)
            Telas.SetFocus
    
        Case 6356  'Tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174581)

     End Select
        
     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    'Limpa a Tela
    Call Limpa_Tela(Tela)
    Call Limpa_Tela_Local

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colModulo As New Collection
Dim vModulo As Variant
Dim sModulo As String
Dim colTela As New Collection
Dim vTela As Variant
Dim objTela As New ClassDicTela
Dim iIndice As Integer

On Error GoTo Erro_Tela_Form_Load

    Me.HelpContextID = IDH_TELA
    
    'Lê Módulos existentes no BD
    lErro = Modulos_Le(colModulo)
    If lErro Then Error 6333
    
    'Preenche List da ComboBox Modulo
    For Each vModulo In colModulo
        Modulo.AddItem (vModulo)
    Next
    
   'Se há uma Tela selecionada
    If Len(gsTela) > 0 Then
        
        'Lê nome do Módulo que contém a Tela
        lErro = Modulo_Le_Tela(gsTela, sModulo)
        If lErro = 6342 Then Error 6334  'Não há módulo contendo Tela
        If lErro Then Error 6335
    
        'Seleciona sModulo na ComboBox Modulo
        Call ListBox_Select(sModulo, Modulo)
        
        'Preenche sigla em objTela
        objTela.sNome = gsTela
        
        'Lê dados da Tela
        lErro = Tela_Le(objTela)
        If lErro = 6350 Then Error 6337 'Tela não cadastrada
        If lErro Then Error 6338
             
        'Coloca dados da Tela na Tela
        Nome.Text = objTela.sNome
        Descricao.Text = objTela.sDescricao
        ProjetoOrig.Caption = objTela.sProjeto_Original
        ClasseOrig.Caption = objTela.sClasse_Original
        ProjetoCust.Text = objTela.sProjeto_Customizado
        ClasseCust.Text = objTela.sClasse_Customizada
        
        'Só permite editar descrição quando for tela de usuário
        If InStr(objTela.sNome, "_USU_") > 0 Then
            Descricao.Locked = False
        Else
            Descricao.Locked = True
        End If
           
        'Seleciona Tela na ListBox Telas
        Call ListBox_Select(objTela.sNome, Telas)
            
        gsTela = ""
        
    Else
    
        'Seleciona o primeiro Módulo na ComboBox Modulo
        Modulo.ListIndex = 0
    
    End If
                 
    lErro_Chama_Tela = SUCESSO
                 
    Exit Sub
    
Erro_Tela_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err
            
        Case 6333, 6335, 6336, 6338, 6339  'Tratado na rotina chamada
        
        Case 6334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_TELA_INEXISTENTE", Err, gsTela)
            
        Case 6337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_CADASTRADA", Err, gsTela)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174582)

    End Select
    
    Exit Sub

End Sub

Private Sub Modulo_Click()

Dim lErro As Long
Dim colTela As New Collection
Dim vTela As Variant

On Error GoTo Erro_Modulo_Click

    'Lê siglas de Telas contidas no Módulo
    lErro = Telas_Le_NomeModulo(Modulo.Text, colTela)
    If lErro Then Error 6352
    
    'Limpa ListBox Telas
    Telas.Clear
    
    'Preenche ListBox Telas
    For Each vTela In colTela
        Telas.AddItem (vTela)
    Next
    
    'Limpa a Tela
    Call Limpa_Tela(Tela)
    Call Limpa_Tela_Local
    
    Exit Sub
    
Erro_Modulo_Click:

    Select Case Err
            
        Case 6352  'Tratado na rotina chamada
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174583)

    End Select
    
    Exit Sub

End Sub

Private Sub Telas_DblClick()

Dim lErro As Long
Dim objTela As New ClassDicTela

On Error GoTo Erro_Telas_DblClick
        
    'Preenche sigla em objTela
    objTela.sNome = Telas.Text
    
    'Lê dados da Tela
    lErro = Tela_Le(objTela)
    If lErro = 6350 Then Error 6353 'Tela não cadastrada
    If lErro Then Error 6354
         
    'Coloca dados da Tela na Tela
    Nome.Text = objTela.sNome
    Descricao.Text = objTela.sDescricao
    ProjetoOrig.Caption = objTela.sProjeto_Original
    ClasseOrig.Caption = objTela.sClasse_Original
    ProjetoCust.Text = objTela.sProjeto_Customizado
    ClasseCust.Text = objTela.sClasse_Customizada
    
    'Só permite editar descrição quando for tela de usuário
    If InStr(objTela.sNome, "_USU_") > 0 Then
        Descricao.Locked = False
    Else
        Descricao.Locked = True
    End If
                         
    Exit Sub
    
Erro_Telas_DblClick:

    Select Case Err
            
        Case 6353
        lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_CADASTRADA", Err, objTela.sNome)
        
        Case 6354  'Tratado na rotina chamada
                 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174584)

    End Select
    
    Exit Sub

End Sub

Private Sub TelaxGrupo_Click()

    'Preenche gsTela
    gsTela = Nome.Text
    
    'Exibe a tela TelaGrupo
    TelaGrupo.Show

End Sub
