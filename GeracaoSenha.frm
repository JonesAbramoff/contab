VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form GeracaoSenha 
   Caption         =   "Geração de Senha"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4200
      Index           =   1
      Left            =   195
      TabIndex        =   1
      Top             =   720
      Width           =   9135
      Begin VB.Frame Frame6 
         Caption         =   "Cliente"
         Height          =   2205
         Left            =   60
         TabIndex        =   14
         Top             =   315
         Width           =   6240
         Begin MSMask.MaskEdBox Serie 
            Height          =   315
            Left            =   1710
            TabIndex        =   15
            Top             =   1255
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CGC 
            Height          =   315
            Left            =   1725
            TabIndex        =   40
            Top             =   1710
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nº Série:"
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
            Left            =   840
            TabIndex        =   21
            Top             =   1330
            Width           =   780
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
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
            Left            =   1185
            TabIndex        =   20
            Top             =   1785
            Width           =   450
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
            Left            =   1050
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   19
            Top             =   420
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nome Reduzido:"
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
            Left            =   225
            TabIndex        =   18
            Top             =   875
            Width           =   1410
         End
         Begin VB.Label Nome 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1725
            TabIndex        =   17
            Top             =   345
            Width           =   4320
         End
         Begin VB.Label NomeReduzido 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1725
            TabIndex        =   16
            Top             =   800
            Width           =   3180
         End
      End
      Begin VB.ListBox Clientes 
         Height          =   3570
         Left            =   6495
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   390
         Width           =   2580
      End
      Begin VB.Frame Frame3 
         Caption         =   "Limites"
         Height          =   1365
         Left            =   45
         TabIndex        =   3
         Top             =   2595
         Width           =   6240
         Begin MSMask.MaskEdBox LimiteEmpresas 
            Height          =   315
            Left            =   1620
            TabIndex        =   4
            Top             =   330
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LimiteFiliais 
            Height          =   315
            Left            =   1620
            TabIndex        =   5
            Top             =   810
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LimiteLogs 
            Height          =   315
            Left            =   4050
            TabIndex        =   6
            Top             =   300
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownValidade 
            Height          =   300
            Left            =   5145
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   825
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ValidadeAte 
            Height          =   300
            Left            =   4080
            TabIndex        =   8
            Top             =   825
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Filiais:"
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
            Left            =   1035
            TabIndex        =   12
            Top             =   870
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Logs:"
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
            Left            =   3525
            TabIndex        =   11
            Top             =   375
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Validade Até:"
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
            Left            =   2835
            TabIndex        =   10
            Top             =   870
            Width           =   1185
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Empresas:"
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
            Left            =   690
            TabIndex        =   9
            Top             =   375
            Width           =   915
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clientes"
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
         Left            =   6480
         TabIndex        =   22
         Top             =   105
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4200
      Index           =   2
      Left            =   195
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame Frame2 
         Caption         =   "Senha"
         Height          =   1350
         Left            =   390
         TabIndex        =   27
         Top             =   2655
         Width           =   7590
         Begin VB.CommandButton BotaoGeraSenha 
            Caption         =   "Gera Senha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   435
            TabIndex        =   34
            Top             =   705
            Width           =   1410
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Grave a senha após transmissão para o Cliente."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   495
            TabIndex        =   41
            Top             =   300
            Width           =   5010
         End
         Begin VB.Label Senha 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   6495
            TabIndex        =   33
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Senha 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   5655
            TabIndex        =   32
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Senha 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   4815
            TabIndex        =   31
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Senha 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   3945
            TabIndex        =   30
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Senha 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   3090
            TabIndex        =   29
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Senha:"
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
            Left            =   2355
            TabIndex        =   28
            Top             =   840
            Width           =   645
         End
      End
      Begin VB.ListBox Modulos 
         Height          =   2085
         Left            =   420
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   495
         Width           =   2940
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   3840
         Picture         =   "GeracaoSenha.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1950
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   3840
         Picture         =   "GeracaoSenha.frx":11E2
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Módulos Liberados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   420
         TabIndex        =   26
         Top             =   210
         Width           =   1605
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7215
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "GeracaoSenha.frx":21FC
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "GeracaoSenha.frx":2386
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "GeracaoSenha.frx":2504
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "GeracaoSenha.frx":2A36
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4725
      Left            =   30
      TabIndex        =   0
      Top             =   345
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   8334
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Limites"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Módulos / Senha"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "GeracaoSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjCliente As ClassCliente
Dim iFrameAtual As Integer
Dim giSenhaGerada As Integer
Dim iLimiteAlterado As Integer
Dim iSenhaGerada As Integer

Private Sub BotaoDesmarcarTodos_Click()

Dim iIndice As Integer

    'Desmarca todos os modulo menos o ADM
    For iIndice = 1 To Modulos.ListCount - 1
        Modulos.Selected(iIndice) = False
    Next
    
    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim lCliente As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se o cliente foi informado
    If Len(Trim(Nome.Caption)) = 0 Then Error 62461
    
    'Recolhe o cliente
    For iIndice = 0 To Clientes.ListCount - 1
        If Clientes.List(iIndice) = NomeReduzido.Caption Then
            lCliente = Clientes.ItemData(iIndice)
            Exit For
        End If
    Next
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CLIENTESLIMITES", lCliente)
    If vbMsgRes = vbNo Then Exit Sub
   
    'Exclui o cliente
    lErro = ClientesLimites_Exclui(lCliente)
    If lErro <> SUCESSO Then Error 62462

    iLimiteAlterado = 0
    iSenhaGerada = False

    'Limpa a tela de geração de senha
    Call Limpa_Tela_GeracaoSenha

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err
    
        Case 62461
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_INFORMADO", Err)
            
        Case 62462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 161534)
            
    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim vbMsgRes As VbMsgBoxResult

    'verifica se os dados dos limites do cliente foram alterados
    'e não foi gerada uma senha
    If iLimiteAlterado = REGISTRO_ALTERADO Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LIMITES_ALTERADOS")
        If vbMsgRes = vbNo Then Exit Sub
    End If
    
    'verifica se uma senha foi gerada e não gravada
    If iSenhaGerada = True Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SENHA_GERADA")
        If vbMsgRes = vbNo Then Exit Sub
    End If
    
    'Fecha a tela
    Unload Me
    
    Exit Sub

End Sub

Private Sub BotaoGeraSenha_Click()

Dim lErro As Long
Dim sSenha As String
Dim colSiglasLib As New Collection
Dim iIndice As Integer
Dim sSigla As String
Dim iPosSeparador As Integer

On Error GoTo Erro_BotaoGeraSenha_Click

    'Verifica o preenchimento dos cpos obrigatórios
    If Len(Trim(Nome.Caption)) = 0 Then Error 62463
    If Len(Trim(Serie.Text)) = 0 Then Error 62464
    If Len(Trim(CGC.Text)) = 0 Then Error 62465
    If Len(Trim(LimiteEmpresas.Text)) = 0 Then Error 62466
    If Len(Trim(LimiteLogs.Text)) = 0 Then Error 62467
    If Len(Trim(LimiteFiliais.Text)) = 0 Then Error 62468
    If Len(Trim(ValidadeAte.ClipText)) = 0 Then Error 62469
       
    'Recolhe os módulos selecionados
    For iIndice = 0 To Modulos.ListCount - 1
        If Modulos.Selected(iIndice) Then
            iPosSeparador = InStr(Modulos.List(iIndice), SEPARADOR)
            sSigla = Mid(Modulos.List(iIndice), 1, iPosSeparador - 1)
            colSiglasLib.Add sSigla
        End If
    Next
    
    'Gear a senha com os dados da tela
    lErro = Senha_Empresa_Gera(CGC.ClipText, Nome.Caption, LimiteLogs, LimiteEmpresas, LimiteFiliais, colSiglasLib, CDate(ValidadeAte), sSenha)
    If lErro <> SUCESSO Then Error 62470
    
    'Coloca a senha na tela
    For iIndice = Senha.LBound To Senha.UBound
        Senha(iIndice).Caption = Mid(sSenha, (iIndice * 5) + 1, 5)
    Next
    
    'Indica que a senha foi gerada
    iSenhaGerada = True
    iLimiteAlterado = 0
    
    Exit Sub
    
Erro_BotaoGeraSenha_Click:

    Select Case Err
        
        Case 62463
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_INFORMADO", Err)
        
        Case 62464
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_SERIE_NAO_PREENCHIDO", Err)
        
        Case 62465
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INFORMADO", Err)
        
        Case 62466
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEEMPRESAS_NAO_INFORMADO", Err)
        
        Case 62467
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITELOGS_NAO_INFORMADO", Err)
        
        Case 62468
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEFILIAIS_NAO_INFORMADO", Err)

        Case 62469
            Call Rotina_Erro(vbOKOnly, "ERRO_VALIDADEATE_NAO_INFORMADA", Err)
        
        Case 62470

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 161535)
    
    End Select
    
    Exit Sub

End Sub

Private Function Gravar_Registro() As Long

Dim lErro As Long
Dim objClientesLimites As New ClassClientesLimites

On Error GoTo Erro_Gravar_Registro

    'verifica o preenchimento dos cpos obrigatórios
    If Len(Trim(Nome.Caption)) = 0 Then Error 62471
    If Len(Trim(Serie.Text)) = 0 Then Error 62472
    If Len(Trim(CGC.Text)) = 0 Then Error 62473
    If Len(Trim(LimiteEmpresas.Text)) = 0 Then Error 62474
    If Len(Trim(LimiteLogs.Text)) = 0 Then Error 62475
    If Len(Trim(LimiteFiliais.Text)) = 0 Then Error 62476
    If Len(Trim(ValidadeAte.ClipText)) = 0 Then Error 62477
    If Len(Trim(Senha(0).Caption)) = 0 Then Error 62478
    
    'recolhe os dados da tela
    lErro = Move_Tela_Memoria(objClientesLimites)
    If lErro <> SUCESSO Then Error 62479
    
    'Grava os Limites do cliente
    lErro = ClientesLimites_Grava(objClientesLimites)
    If lErro <> SUCESSO Then Error 62480
    
    'Limpa a tela
    Call Limpa_Tela_GeracaoSenha
        
    'Zera flags de alteração
    iLimiteAlterado = 0
    iSenhaGerada = False
    
    Gravar_Registro = SUCESSO
       
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err
    
    Select Case Err
    
        Case 62471
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_INFORMADO", Err)
        
        Case 62472
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_SERIE_NAO_PREENCHIDO", Err)
        
        Case 62473
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INFORMADO", Err)
        
        Case 62474
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEEMPRESAS_NAO_INFORMADO", Err)
        
        Case 62475
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITELOGS_NAO_INFORMADO", Err)
        
        Case 62476
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEFILIAIS_NAO_INFORMADO", Err)

        Case 62477
            Call Rotina_Erro(vbOKOnly, "ERRO_VALIDADEATE_NAO_INFORMADA", Err)
        
        Case 62478
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_GERADA", Err)
        
        Case 62479, 62480
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 161536)
    
    End Select
    
    Exit Function
        
End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 62534
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 62534
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 161537)
            
    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim vbMsgRes As VbMsgBoxResult

    'Verifica se os limites foram alterados e não foi gerada senha
    If iLimiteAlterado = REGISTRO_ALTERADO Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LIMITES_ALTERADOS")
        If vbMsgRes = vbNo Then Exit Sub
    End If
    
    'Verifica se uma senha foi gerada e não gravada
    If iSenhaGerada = True Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SENHA_GERADA")
        If vbMsgRes = vbNo Then Exit Sub
    End If
    
    'Limpa a tela
    Call Limpa_Tela_GeracaoSenha
    
    'Zera flas de alteração
    iLimiteAlterado = 0
    iSenhaGerada = False
    
    Exit Sub
    
End Sub

Private Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer
    
    'Seleciona todos os módulos
    For iIndice = 0 To Modulos.ListCount - 1
        Modulos.Selected(iIndice) = True
    Next
    
    Exit Sub

End Sub

Private Sub CGC_LostFocus()

Dim lErro As Long

On Error GoTo Erro_CGC_LostFocus
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGC.Text))
        
        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 62481
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else
                
            Error 62482

    End Select

    Exit Sub

Erro_CGC_LostFocus:

    Select Case Err

        Case 62481

        Case 62482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161538)

    End Select

    CGC.SetFocus

    Exit Sub

End Sub

Private Sub Clientes_DblClick()

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Clientes_DblClick

    'verifica se tem algum cliente selecionado
    If Clientes.ListIndex = -1 Then Exit Sub
    
    'pega o código do cliente selecionado
    objCliente.lCodigo = Clientes.ItemData(Clientes.ListIndex)

    'Lê o cliente
    lErro = CF("Cliente_Le",objCliente)
    If lErro <> SUCESSO And lErro <> 19062 Then Error 62483
    If lErro <> SUCESSO Then Error 62484 'Não encontrou
    
    'Traz os dados do cliente para a tela
    lErro = Traz_DadosCliente_Tela(objCliente)
    If lErro <> SUCESSO Then Error 62485
    
    'Zera flags de alteração
    iLimiteAlterado = 0
    iSenhaGerada = 0
    
    Exit Sub

Erro_Clientes_DblClick:

    Select Case Err
        
        Case 62483, 62485
        
        Case 62484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161539)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Form_Load()
    
Dim lErro As Long

On Error GoTo Erro_Form_Load

    'carrega a list de Clientes
    lErro = Carrega_Clientes()
    If lErro <> SUCESSO Then Error 62486
    
    'Carrega a list de módulos
    lErro = Carrega_Modulos()
    If lErro <> SUCESSO Then Error 62487
    
    iFrameAtual = 1
    iLimiteAlterado = 0
    iSenhaGerada = False
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    Select Case Err

        Case 62486, 62487

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161540)

    End Select

    Exit Sub

End Sub

Function Carrega_Clientes() As Long

Dim lErro  As Long
Dim colCodigoNome As New AdmCollCodigoNome
Dim objCodigoNome As AdmlCodigoNome

On Error GoTo Erro_Carrega_Clientes

    'Lê Códigos e NomesReduzidos da tabela de Clientes e devolve na coleção
    lErro = CF("LCod_Nomes_Le","Clientes", "Codigo", "NomeReduzido", 50, colCodigoNome)
    If lErro <> SUCESSO Then Error 62488

    'Preenche a ListBox ClientesList com os objetos da coleção
    For Each objCodigoNome In colCodigoNome
        Clientes.AddItem objCodigoNome.sNome
        Clientes.ItemData(Clientes.NewIndex) = objCodigoNome.lCodigo
    Next

    Carrega_Clientes = SUCESSO
    
    Exit Function
    
Erro_Carrega_Clientes:

    Carrega_Clientes = Err
    
    Select Case Err
    
        Case 62488
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161541)
            
    End Select

    Exit Function

End Function

Function Carrega_Modulos() As Long

Dim lErro  As Long
Dim objModulos As AdmModulo
Dim colModulos As New Collection

On Error GoTo Erro_Carrega_Modulos

    'Lê Códigos e NomesReduzidos da tabela de Clientes e devolve na coleção
    lErro = CF("Modulos_Le_Todos",colModulos)
    If lErro <> SUCESSO Then Error 62489

    'Preenche a ListBox ClientesList com os objetos da coleção
    For Each objModulos In colModulos
        Modulos.AddItem objModulos.sSigla & SEPARADOR & objModulos.sDescricao
        If objModulos.sSigla = MODULO_ADM Then Modulos.Selected(Modulos.NewIndex) = True
    Next

    Carrega_Modulos = SUCESSO
    
    Exit Function
    
Erro_Carrega_Modulos:

    Carrega_Modulos = Err
    
    Select Case Err
    
        Case 62489
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161542)
            
    End Select

    Exit Function

End Function

Private Sub LimiteEmpresas_Change()

    iLimiteAlterado = REGISTRO_ALTERADO
    iSenhaGerada = False
    
    Call Limpa_Senha

End Sub
Private Sub LimiteEmpresas_Validate(Cancel As Boolean)

On Error GoTo Erro_LimiteEmpresas_Validate

    If Len(Trim(LimiteEmpresas.Text)) = 0 Then Exit Sub
    
    If StrParaInt(LimiteEmpresas) = 0 Then Error 62533
    
    If Len(Trim(LimiteFiliais.Text)) = 0 Then Exit Sub
    
    If StrParaInt(LimiteFiliais) > StrParaInt(LimiteEmpresas.Text) Then Error 62532
    
    Cancel = False
    
    Exit Sub
        
Erro_LimiteEmpresas_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 62532
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEEMPRESAS_MAIOR_LIMITEFILIAIS", Err)
        
        Case 62533
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, LimiteEmpresas.Text)
        
    End Select

    Exit Sub

End Sub
Private Sub LimiteFiliais_Change()

    iLimiteAlterado = REGISTRO_ALTERADO
    iSenhaGerada = False
    Call Limpa_Senha

End Sub


Private Sub LimiteFiliais_Validate(Cancel As Boolean)

On Error GoTo Erro_LimiteFiliais_Validate

    If Len(Trim(LimiteFiliais.Text)) = 0 Then Exit Sub
    
    If StrParaInt(LimiteFiliais) = 0 Then Error 62533
    
    If Len(Trim(LimiteEmpresas.Text)) = 0 Then Exit Sub
    
    If StrParaInt(LimiteEmpresas) > StrParaInt(LimiteFiliais) Then Error 62532
    
    Cancel = False
    
    Exit Sub
        
Erro_LimiteFiliais_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 62532
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEEMPRESAS_MAIOR_LIMITEFILIAIS", Err)
        
        Case 62533
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, LimiteFiliais.Text)
        
    End Select

    Exit Sub

End Sub

Private Sub LimiteLogs_Change()
    
    iLimiteAlterado = REGISTRO_ALTERADO
    iSenhaGerada = False
    Call Limpa_Senha

End Sub

Private Sub Modulos_ItemCheck(Item As Integer)
        
    If Item = 0 Then Modulos.Selected(0) = True
    
    iLimiteAlterado = REGISTRO_ALTERADO
    iSenhaGerada = False
    
    Call Limpa_Senha
    
End Sub

Public Sub Opcao_Click()
    
    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me, 0) <> SUCESSO Then Exit Sub
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

    End If

End Sub

Private Sub ValidadeAte_Change()

    iLimiteAlterado = REGISTRO_ALTERADO
    iSenhaGerada = 0
    
    Call Limpa_Senha

End Sub

Private Sub ValidadeAte_LostFocus()

Dim lErro As Long

On Error GoTo Erro_ValidadeAte_LostFocus

    'Verifica se a data de emissao foi digitada
    If Len(Trim(ValidadeAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(ValidadeAte.Text)
    If lErro <> SUCESSO Then Error 62490
    
    If Date > CDate(ValidadeAte) Then Error 62531

    Exit Sub

Erro_ValidadeAte_LostFocus:

    Select Case Err

        Case 62490
            ValidadeAte.SetFocus
        
        Case 62531
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVALIDADE_MENOR", Err, ValidadeAte.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161543)

    End Select

    Exit Sub


End Sub
Public Sub UpDownValidade_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownValidade_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(ValidadeAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 62491

    Exit Sub

Erro_UpDownValidade_DownClick:

    Select Case Err

        Case 62491

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161544)

    End Select

    Exit Sub

End Sub

Public Sub UpDownValidade_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownValidade_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(ValidadeAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 62492

    Exit Sub

Erro_UpDownValidade_UpClick:

    Select Case Err

        Case 62492

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161545)

    End Select

    Exit Sub

End Sub

Function Traz_DadosCliente_Tela(objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objClientesLimites As New ClassClientesLimites
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iPosicao As Integer

On Error GoTo Erro_Traz_DadosCliente_Tela

    Call Limpa_Tela_GeracaoSenha
    
    'Coloca o nome e o nome reduzido na tela
    Nome.Caption = objCliente.sRazaoSocial
    NomeReduzido.Caption = objCliente.sNomeReduzido
        
    objClientesLimites.lCodCliente = objCliente.lCodigo
    
    'Busca no BD os Limites do cliente
    lErro = ClientesLimites_Le(objClientesLimites)
    If lErro <> SUCESSO And lErro <> 62498 Then Error 62493
    'Se encontrou
    If lErro = SUCESSO Then
        'Coloca os dados dos limites na tela
        LimiteEmpresas.Text = objClientesLimites.iLimiteEmpresas
        LimiteFiliais.Text = objClientesLimites.iLimiteFiliais
        LimiteLogs.Text = objClientesLimites.iLimiteLogs
        ValidadeAte.PromptInclude = False
        ValidadeAte.Text = objClientesLimites.dtValidadeAte
        ValidadeAte.PromptInclude = True
        
        For iIndice = Senha.LBound To Senha.UBound
            Senha(iIndice).Caption = Mid(objClientesLimites.sSenha, (iIndice * 5) + 1, 5)
        Next
        
        For iIndice1 = 1 To objClientesLimites.colSiglasModulosLib.Count
            iPosicao = 0
            For iIndice = 0 To Modulos.ListCount - 1
                iPosicao = InStr(Modulos.List(iIndice), objClientesLimites.colSiglasModulosLib(iIndice1) & SEPARADOR)
                If iPosicao > 0 Then
                    Modulos.Selected(iIndice) = True
                    Exit For
                End If
            Next
        Next
    
    End If
        
    Traz_DadosCliente_Tela = SUCESSO
    
    Exit Function

Erro_Traz_DadosCliente_Tela:

    Traz_DadosCliente_Tela = Err
    
    Select Case Err
        
        Case 62493
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161546)
            
    End Select
    
    Exit Function
    
End Function

Function ModuloCliente_Le(lCodCliente As Long, colSiglasModulos As Collection) As Long
'Lê os modulos liberados para o cliente passado

Dim lErro As Long
Dim lComando As Long
Dim sSigla As String

On Error GoTo Erro_ModuloCliente_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62494
    
    sSigla = String(STRING_MODULO_SIGLA, 0)
    
    'Busca os módulos lib para o cliente passado
    lErro = Comando_Executar(lComando, "SELECT SiglaModulo FROM ModuloCliente WHERE CodCliente = ?", sSigla, lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 62495
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62496
    
    'Para cada módulo encontrado
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Adicona o módulo na coleção de módulos
        colSiglasModulos.Add sSigla
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62497
    
    Loop
         
    Call Comando_Fechar(lComando)
    
    ModuloCliente_Le = SUCESSO
    
    Exit Function

Erro_ModuloCliente_Le:

    ModuloCliente_Le = Err
    
    Select Case Err
    
        Case 62494
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62495, 62496, 62497
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOCLIENTE", Err, lCodCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161547)
    
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ClientesLimites_Le(objClientesLimites As ClassClientesLimites) As Long
'Lê os limites cadastrados para o cliente passado

Dim lErro As Long
Dim lComando As Long
Dim tClientesLimites As typeClientesLimites

On Error GoTo Erro_ClientesLimites_Le
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62497
    
    With tClientesLimites
    
        .sSenha = String(STRING_DICCONFIG_SENHA, 0)
        .sSerie = String(STRING_DICCONFIG_SERIE, 0)
        .sVersao = String(STRING_MODULO_VERSAO, 0)
        
        'Busca os limites do cliente
        lErro = Comando_Executar(lComando, "SELECT CodFilial,Serie,TipoVersao,LimiteLogs,LimiteFiliais,LimiteEmpresas,Senha,DataSenha,ValidadeAte,Versao FROM ClientesLimites WHERE CodCliente = ?", .iCodFilial, .sSerie, .iTipoVersao, .iLimiteLogs, .iLimiteFiliais, .iLimiteEmpresas, .sSenha, .dtDataSenha, .dtValidadeAte, .sVersao, objClientesLimites.lCodCliente)
        If lErro <> AD_SQL_SUCESSO Then Error 62530
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62500
        If lErro <> AD_SQL_SUCESSO Then Error 62498 'Não encontrou
        
        'Preenche o obj com os dados lidos
        objClientesLimites.dtDataSenha = .dtDataSenha
        objClientesLimites.dtValidadeAte = .dtValidadeAte
        objClientesLimites.iCodFilial = .iCodFilial
        objClientesLimites.iLimiteEmpresas = .iLimiteEmpresas
        objClientesLimites.iLimiteFiliais = .iLimiteFiliais
        objClientesLimites.iLimiteLogs = .iLimiteLogs
        objClientesLimites.iTipoVersao = .iTipoVersao
        objClientesLimites.sSenha = .sSenha
        objClientesLimites.sSerie = .sSerie
        objClientesLimites.sVersao = .sVersao
        
    End With
    
    'Lê os módulos liberados
    lErro = ModuloCliente_Le(objClientesLimites.lCodCliente, objClientesLimites.colSiglasModulosLib)
    If lErro <> SUCESSO Then Error 62499
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ClientesLimites_Le = SUCESSO
        
    Exit Function
    
Erro_ClientesLimites_Le:

    ClientesLimites_Le = Err
    
    Select Case Err
        
        Case 62497
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62498, 62499
        
        Case 62500, 62530
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESLIMITES", Err, objClientesLimites.lCodCliente)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161548)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Move_Tela_Memoria(objClientesLimites As ClassClientesLimites) As Long
'Recolhe os dados da tela para o objClientesLimites

Dim lErro As Long
Dim iCodFilial As Integer
Dim iIndice As Integer
Dim iPosSeparador As Integer
Dim sSigla As String

On Error GoTo Erro_Move_Tela_Memoria
    
    'Recolhe os dados da tela
    objClientesLimites.dtDataSenha = Date
    objClientesLimites.dtValidadeAte = StrParaDate(ValidadeAte.Text)
    objClientesLimites.iLimiteEmpresas = StrParaInt(LimiteEmpresas.Text)
    objClientesLimites.iLimiteFiliais = StrParaInt(LimiteFiliais.Text)
    objClientesLimites.iLimiteLogs = StrParaInt(LimiteLogs.Text)
    objClientesLimites.sSenha = Senha(0) & Senha(1) & Senha(2) & Senha(3) & Senha(4)
    objClientesLimites.sSerie = Serie.Text
    objClientesLimites.sVersao = ""
    objClientesLimites.sCgc = CGC.ClipText

    'Busca o cliente na list
    For iIndice = 0 To Clientes.ListCount - 1
        If Clientes.List(iIndice) = NomeReduzido.Caption Then
            objClientesLimites.lCodCliente = Clientes.ItemData(iIndice)
            Exit For
        End If
    Next
    
    'recolhe os módulos liberados
    For iIndice = 0 To Modulos.ListCount - 1
        If Modulos.Selected(iIndice) Then
            iPosSeparador = InStr(Modulos.List(iIndice), SEPARADOR)
            sSigla = Mid(Modulos.List(iIndice), 1, iPosSeparador - 1)
            objClientesLimites.colSiglasModulosLib.Add sSigla
        End If
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161549)

    End Select

    Exit Function

End Function

Function ClientesLimites_Grava(objClientesLimites As ClassClientesLimites) As Long
'Grava ou atualiza os limites do cliente passado

Dim lErro As Long
Dim iCodFilial As Integer
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim sSerie As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ClientesLimites_Grava
    
    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 62501

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 62502
    Next

    'Busca no BD alguma filial do clienet com o CGC informado
    lErro = Comando_Executar(alComando(0), "SELECT CodFilial FROM FiliaisClientes WHERE CodCliente = ? AND CGC = ?", iCodFilial, objClientesLimites.lCodCliente, objClientesLimites.sCgc)
    If lErro <> AD_SQL_SUCESSO Then Error 62503
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62504
    If lErro <> AD_SQL_SUCESSO Then Error 62505 'não encontrou
    
    'Guarda o código da filial
    objClientesLimites.iCodFilial = iCodFilial
    
    'Faz o lock na Filial do cliente
    lErro = CF("FilialCliente_Lock",objClientesLimites.lCodCliente, objClientesLimites.iCodFilial)
    If lErro <> SUCESSO Then Error 62506
    
    sSerie = String(STRING_DICCONFIG_SERIE, 0)
    
    'Busca no BD os Limites do cliente
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodFilial,Serie FROM ClientesLimites WHERE CodCliente = ?", 0, iCodFilial, sSerie, objClientesLimites.lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 62507
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62508
    'Se encontrou
    If lErro = AD_SQL_SUCESSO Then
        
        If sSerie <> objClientesLimites.sSerie Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_SERIE", sSerie, objClientesLimites.sSerie)
            If vbMsgRes = vbNo Then Error 62509
        End If
        
        'Atualiza os limites fornecidos
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE ClientesLimites SET CodFilial = ?, Serie =?, LimiteLogs=?, LimiteFiliais=?, LimiteEmpresas=?, Senha=?, DataSenha=?, ValidadeAte=?", alComando(0), objClientesLimites.iCodFilial, objClientesLimites.sSerie, objClientesLimites.iLimiteLogs, objClientesLimites.iLimiteFiliais, objClientesLimites.iLimiteEmpresas, objClientesLimites.sSenha, objClientesLimites.dtDataSenha, objClientesLimites.dtValidadeAte)
        If lErro <> AD_SQL_SUCESSO Then Error 62510
    'Senão
    Else
        'Insere
        lErro = Comando_Executar(alComando(0), "INSERT INTO ClientesLimites (CodCliente,CodFilial,Serie,LimiteLogs,LimiteFiliais,LimiteEmpresas,Senha,DataSenha,ValidadeAte,TipoVersao,Versao) VALUES (?,?,?,?,?,?,?,?,?,?,?) ", objClientesLimites.lCodCliente, objClientesLimites.iCodFilial, objClientesLimites.sSerie, objClientesLimites.iLimiteLogs, objClientesLimites.iLimiteFiliais, objClientesLimites.iLimiteEmpresas, objClientesLimites.sSenha, objClientesLimites.dtDataSenha, objClientesLimites.dtValidadeAte, objClientesLimites.iTipoVersao, objClientesLimites.sVersao)
        If lErro <> AD_SQL_SUCESSO Then Error 62511
    
    End If
    
    'Exclui os módulos liberados do cliente
    lErro = ModuloCliente_Exclui(objClientesLimites.lCodCliente)
    If lErro <> SUCESSO Then Error 62512
    
    'Inclui os módulos liberados passados no obj
    lErro = ModuloCliente_Inclui(objClientesLimites.lCodCliente, objClientesLimites.colSiglasModulosLib)
    If lErro <> SUCESSO Then Error 62513
    
    'Confirma as alteração no BD
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 62514
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ClientesLimites_Grava = SUCESSO
    
    Exit Function

Erro_ClientesLimites_Grava:

    ClientesLimites_Grava = Err
    
    Select Case Err
    
        Case 62501
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 62502
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62503, 62504
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES2", Err)
        
        Case 62505
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_CGC_NAO_ENCONTRADA", Err, objClientesLimites.sCgc)
        
        Case 62506, 62512, 62513, 62509
        
        Case 62507, 62508
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESLIMITES", Err, objClientesLimites.lCodCliente)
        
        Case 62510
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CLIENTESLIMITES", Err, objClientesLimites.lCodCliente)

        Case 62511
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CLIENTESLIMITES", Err, objClientesLimites.lCodCliente)
        
        Case 62514
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161550)

    End Select
    
    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ModuloCliente_Exclui(lCodCliente As Long) As Long
'Exclui os módulos liberados do cliente passadoi

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice  As Integer
Dim lCodigo As Long

On Error GoTo Erro_ModuloCliente_Exclui:

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 62515
    Next
    
    'BUsca os módulos liberados para o cliente
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodCliente FROM ModuloCliente WHERE Codcliente = ?", 0, lCodigo, lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 62516
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62517
    'Para cada módulo encontrado
    Do While lErro = AD_SQL_SUCESSO
        
        'Exclui o módulo do BD
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ModuloCliente", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 62518
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62519

    Loop

    'Fecha os comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ModuloCliente_Exclui = SUCESSO
    
    Exit Function

Erro_ModuloCliente_Exclui:

    ModuloCliente_Exclui = Err
    
    Select Case Err
    
        Case 62515
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62516, 62517, 62519
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOCLIENTE", Err, lCodCliente)
        
        Case 62518
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MODULOCLIENTE", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161551)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next


    Exit Function

End Function

Function ModuloCliente_Inclui(lCodCliente As Long, colSiglasModulosLib As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer

On Error GoTo Erro_ModuloCliente_Inclui
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62520
    
    'Para cada módulo passado
    For iIndice = 1 To colSiglasModulosLib.Count
        'Inclui o módulo na tabela ModuloCliente
        lErro = Comando_Executar(lComando, "INSERT INTO ModuloCliente (CodCliente,SiglaModulo) VALUES (?,?)", lCodCliente, colSiglasModulosLib(iIndice))
        If lErro <> AD_SQL_SUCESSO Then Error 62521
    
    Next
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ModuloCliente_Inclui = SUCESSO
    
    Exit Function
    
Erro_ModuloCliente_Inclui:

    ModuloCliente_Inclui = Err
    
    Select Case Err
    
        Case 62520
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62521
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MODULOCLIENTE", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161552)
            
    End Select
        
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Private Sub Limpa_Tela_GeracaoSenha()
'Limpa a tela de Geração de Senha

    'Limpa os controle comuns da tela
    Call Limpa_Tela(Me)
    
    'Limapa os controle especiais
    Nome.Caption = ""
    NomeReduzido.Caption = ""
    
    'Deseliciona os módulos na list
    BotaoDesmarcarTodos_Click
        
    Call Limpa_Senha

    Exit Sub

End Sub

Function ClientesLimites_Exclui(lCliente As Long) As Long
'Exclui os limites do sistema do cliente passado

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iCodFilial As Integer
Dim lTransacao As Long

On Error GoTo Erro_ClientesLimites_Exclui

    'Abrir transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 62522

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 62523
    Next

    'Busca os limites do cliente no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodFilial FROM ClientesLimites WHERE CodCliente =? ", 0, iCodFilial, lCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 62524
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62525
    If lErro <> AD_SQL_SUCESSO Then Error 62526 'Não encontrou
    
    'Exclui os módulos liberados para o cliente
    lErro = ModuloCliente_Exclui(lCliente)
    If lErro <> SUCESSO Then Error 62527
    
    'Exclui os limites do cliente
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ClientesLimites ", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 62528

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 62529
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ClientesLimites_Exclui = SUCESSO
    
    Exit Function
        
Erro_ClientesLimites_Exclui:

    ClientesLimites_Exclui = Err
    
    Select Case Err

        Case 62522
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 62523
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 62524, 62525
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTESLIMITES", Err, lCliente)

        Case 62526
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITES_CLIENTE_NAO_CADASTRADOS", Err, lCliente)

        Case 62527

        Case 62528
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTESLIMITES", Err, lCliente)

        Case 62529
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Sub Limpa_Senha()

    'Limpa a senha
    Senha(0) = ""
    Senha(1) = ""
    Senha(2) = ""
    Senha(3) = ""
    Senha(4) = ""

    Exit Sub
    
End Sub
