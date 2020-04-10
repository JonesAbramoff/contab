VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RotinaTela 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rotina"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   960
   ClientWidth     =   9480
   Icon            =   "Rotina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.Frame SSFrame1 
      Caption         =   "Rotina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   420
      TabIndex        =   11
      Top             =   690
      Width           =   5415
      Begin VB.Frame SSFrame2 
         Caption         =   "Customizada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   420
         TabIndex        =   21
         Top             =   2355
         Width           =   4485
         Begin MSMask.MaskEdBox ClasseCust 
            Height          =   315
            Left            =   1470
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
            Left            =   540
            TabIndex        =   25
            Top             =   330
            Width           =   705
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
            Left            =   600
            TabIndex        =   24
            Top             =   765
            Width           =   645
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
         Left            =   420
         TabIndex        =   16
         Top             =   1155
         Width           =   4500
         Begin VB.Label ClasseOrig 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1470
            TabIndex        =   20
            Top             =   690
            Width           =   2730
         End
         Begin VB.Label ProjetoOrig 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1470
            TabIndex        =   19
            Top             =   270
            Width           =   2730
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
            Left            =   600
            TabIndex        =   18
            Top             =   735
            Width           =   645
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
            Left            =   570
            TabIndex        =   17
            Top             =   330
            Width           =   705
         End
      End
      Begin VB.TextBox Sigla 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   330
         Width           =   2730
      End
      Begin VB.TextBox Descricao 
         Height          =   315
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   750
         Width           =   2730
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sigla:"
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
         Left            =   1170
         TabIndex        =   15
         Top             =   360
         Width           =   525
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
         Left            =   735
         TabIndex        =   14
         Top             =   765
         Width           =   945
      End
   End
   Begin VB.Frame SSFrame3 
      Caption         =   "Observação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   390
      TabIndex        =   9
      Top             =   4410
      Width           =   5445
      Begin VB.Label Label8 
         Caption         =   "A customização desta tela é a nível de Empresa. Para customizar a nível de Grupo clique o Botão  Rotina x Grupo."
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   270
         Width           =   4305
      End
   End
   Begin VB.CommandButton RotinaxGrupo 
      Caption         =   "Rotina x Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6885
      TabIndex        =   8
      Top             =   4485
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7545
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   90
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Rotina.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "Rotina.frx":02A4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "Rotina.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Rotinas 
      Height          =   2985
      ItemData        =   "Rotina.frx":0954
      Left            =   6210
      List            =   "Rotina.frx":0956
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1275
      Width           =   3015
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      Left            =   2355
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rotinas"
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
      Left            =   6195
      TabIndex        =   3
      Top             =   1020
      Width           =   675
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
      Left            =   1500
      TabIndex        =   2
      Top             =   255
      Width           =   705
   End
End
Attribute VB_Name = "RotinaTela"
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

    Unload RotinaTela

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objRotina As New ClassDicRotina

On Error GoTo Erro_BotaoGravar_Click

    'Verifica se dados da Rotina foram informados
    If Len(Sigla.Text) = 0 Then Error 6331
    
    objRotina.sSigla = Sigla.Text
    objRotina.sDescricao = Trim(Descricao.Text)
    objRotina.sProjeto_Customizado = Trim(ProjetoCust.Text)
    objRotina.sClasse_Customizada = Trim(ClasseCust.Text)
        
    'grava a Rotina no banco de dados (é um update)
    lErro = Rotina_Grava(objRotina)
    If lErro Then Error 6332
         
    'Limpa a Tela
    Call Limpa_Tela(RotinaTela)
    Call Limpa_Tela_Local
  
Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 6331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ROTINA_NAO_INFORMADA", Err)
            Rotinas.SetFocus
    
        Case 6332  'Tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174308)

     End Select
        
     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    'Limpa a Tela
    Call Limpa_Tela(RotinaTela)
    Call Limpa_Tela_Local

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colModulo As New Collection
Dim vModulo As Variant
Dim sModulo As String
Dim colRotina As New Collection
Dim vRotina As Variant
Dim objRotina As New ClassDicRotina
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Form_Load

    Me.HelpContextID = IDH_ROTINA_TELA
    
    'Lê Módulos existentes no BD
    lErro = Modulos_Le(colModulo)
    If lErro Then Error 6300
    
    'Preenche List da ComboBox Modulo
    For Each vModulo In colModulo
        Modulo.AddItem (vModulo)
    Next
    
   'Se há uma Rotina selecionada
    If Len(gsRotina) > 0 Then
        
        'Lê nome do Módulo que contém a Rotina
        lErro = Modulo_Le_Rotina(gsRotina, sModulo)
        If lErro = 6308 Then Error 6310  'Não há módulo contendo Rotina
        If lErro Then Error 6301

        'Seleciona sModulo na ComboBox Modulo
        Call ListBox_Select(sModulo, Modulo)
        
        'Preenche sigla em objRotina
        objRotina.sSigla = gsRotina
        
        'Lê dados da Rotina
        lErro = Rotina_Le(objRotina)
        If lErro = 6317 Then Error 6319 'Rotina não cadastrada
        If lErro Then Error 6303
             
        'Coloca dados da Rotina na Tela
        Sigla.Text = objRotina.sSigla
        Descricao.Text = objRotina.sDescricao
        ProjetoOrig.Caption = objRotina.sProjeto_Original
        ClasseOrig.Caption = objRotina.sClasse_Original
        ProjetoCust.Text = objRotina.sProjeto_Customizado
        ClasseCust.Text = objRotina.sClasse_Customizada
        
        'Só permite editar descrição quando for rotina de usuário
        If InStr(objRotina.sSigla, "_USU_") > 0 Then
            Descricao.Locked = False
        Else
            Descricao.Locked = True
        End If
           
        'Seleciona Rotina na ListBox Rotinas
        Call ListBox_Select(objRotina.sSigla, Rotinas)
            
        gsRotina = ""
        
    Else
    
        'Seleciona o primeiro Módulo na ComboBox Modulo
        Modulo.ListIndex = 0
    
    End If
                 
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Rotina_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case 6300, 6301, 6302, 6303, 6304  'Tratado na rotina chamada
        
        Case 6310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_ROTINA_INEXISTENTE", Err, gsRotina)
            
        Case 6319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_CADASTRADA", Err, gsRotina)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174309)

    End Select
        
    Exit Sub

End Sub

Private Sub Modulo_Click()

Dim lErro As Long
Dim colRotina As New Collection
Dim vRotina As Variant

On Error GoTo Erro_Modulo_Click

    'Lê siglas de Rotinas contidas no Módulo
    lErro = Rotinas_Le_NomeModulo(Modulo.Text, colRotina)
    If lErro Then Error 6320
    
    'Limpa ListBox Rotinas
    Rotinas.Clear
    
    'Preenche ListBox Rotinas
    For Each vRotina In colRotina
        Rotinas.AddItem (vRotina)
    Next
    
    'Limpa a Tela
    Call Limpa_Tela(RotinaTela)
    Call Limpa_Tela_Local
    
    Exit Sub
    
Erro_Modulo_Click:

    Select Case Err
            
        Case 6320  'Tratado na rotina chamada
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174310)

    End Select
    
    Exit Sub

End Sub

Private Sub Rotinas_DblClick()

Dim lErro As Long
Dim objRotina As New ClassDicRotina

On Error GoTo Erro_Rotinas_DblClick
        
    'Preenche sigla em objRotina
    objRotina.sSigla = Rotinas.Text
    
    'Lê dados da Rotina
    lErro = Rotina_Le(objRotina)
    If lErro = 6317 Then Error 6321 'Rotina não cadastrada
    If lErro Then Error 6322
         
    'Coloca dados da Rotina na Tela
    Sigla.Text = objRotina.sSigla
    Descricao.Text = objRotina.sDescricao
    ProjetoOrig.Caption = objRotina.sProjeto_Original
    ClasseOrig.Caption = objRotina.sClasse_Original
    ProjetoCust.Text = objRotina.sProjeto_Customizado
    ClasseCust.Text = objRotina.sClasse_Customizada
    
    'Só permite editar descrição quando for rotina de usuário
    If InStr(objRotina.sSigla, "_USU_") > 0 Then
        Descricao.Locked = False
    Else
        Descricao.Locked = True
    End If
                         
    Exit Sub
    
Erro_Rotinas_DblClick:

    Select Case Err
            
        Case 6321
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_CADASTRADA", Err, objRotina.sSigla)
        
        Case 6322  'Tratado na rotina chamada
                 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174311)

    End Select
    
    Exit Sub

End Sub

Private Sub RotinaxGrupo_Click()

    'Preenche gsRotina
    gsRotina = Sigla.Text
    
    'Exibe a tela RotinaGrupo
    RotinaGrupo.Show

End Sub
