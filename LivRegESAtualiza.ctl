VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LivRegESAtualizaOcx 
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   5820
   Begin VB.OptionButton OptionForn 
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2625
      TabIndex        =   14
      Top             =   1680
      Width           =   1935
   End
   Begin VB.OptionButton OptionCli 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   255
      TabIndex        =   13
      Top             =   1680
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Frame FrameCliForn 
      Caption         =   "Cliente"
      Height          =   795
      Left            =   135
      TabIndex        =   8
      Top             =   2040
      Width           =   5505
      Begin MSMask.MaskEdBox CliFornDe 
         Height          =   300
         Left            =   690
         TabIndex        =   9
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CliFornAte 
         Height          =   300
         Left            =   3315
         TabIndex        =   10
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelCliFornAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   2895
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   360
         Width           =   360
      End
      Begin VB.Label LabelCliFornDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alterar"
      Height          =   795
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   5535
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   1560
         TabIndex        =   6
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A partir de:"
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
         Left            =   405
         TabIndex        =   7
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.CheckBox LivrosFechados 
      Caption         =   "Altera Livros Fechados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   3
      Top             =   2970
      Width           =   2475
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4020
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   1620
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "LivRegESAtualiza.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "LivRegESAtualiza.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "LivRegESAtualiza.ctx":02D8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "LivRegESAtualizaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1


Dim iAlterado As Integer

Dim iClienteAtual As Integer
Dim iFornecedorAtual As Integer
Const CLIFORN_DE = 1
Const CLIFORN_ATE = 2

Dim iTratamentoAtual As Integer
Const TRATA_CLIENTE = 1
Const TRATA_FORNECEDOR = 2

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Livros de E\S - Atualização dos Dados de Clientes\Fornecedores"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "LivRegESAtualiza"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa eventos de browser
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    iClienteAtual = CLIFORN_DE
    iFornecedorAtual = CLIFORN_DE
    iTratamentoAtual = TRATA_CLIENTE
    iAlterado = REGISTRO_INALTERADO
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182045)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182046)

    End Select

    Exit Function

End Function

Sub Form_Unload(Cancel As Integer)

On Error GoTo Erro_Form_Unload

    Set objEventoCliente = Nothing
    Set objEventoFornecedor = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182047)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182048)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objLivRegESAtualiza As New ClassLivRegESAtualiza

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    If StrParaDate(Data.Text) = DATA_NULA Then gError 182049
    
    lErro = Move_Tela_Memoria(objLivRegESAtualiza)
    If lErro <> SUCESSO Then gError 182050
    
    If objLivRegESAtualiza.lFornecedorAte <> 0 And objLivRegESAtualiza.lFornecedorDe <> 0 Then
        If objLivRegESAtualiza.lFornecedorAte < objLivRegESAtualiza.lFornecedorDe Then gError 182051
    End If

    If objLivRegESAtualiza.lClienteAte <> 0 And objLivRegESAtualiza.lClienteDe <> 0 Then
        If objLivRegESAtualiza.lClienteAte < objLivRegESAtualiza.lClienteDe Then gError 182052
    End If
    
    lErro = CF("LivRegES_Atualiza_Emitentes", objLivRegESAtualiza)
    If lErro <> SUCESSO Then gError 182053

    'Limpa Tela
    Call Limpa_Tela_LivRegES
    
    'fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")
    
    Exit Sub

Erro_BotaoGravar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 182049
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAOPREENCHIDA", gErr)

        Case 182050, 182053
        
        Case 182051
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
        
        Case 182052
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182054)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 182055

    'Limpa a tela
    Call Limpa_Tela_LivRegES

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 182055

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182056)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 182057

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 182057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182058)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 182059

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 182059

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182060)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 182061

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 182061

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182062)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_Preenche(objControle As Object)

Static sNomeReduzidoParte As String
Dim lErro As Long
    
On Error GoTo Erro_Cliente_Preenche
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objControle, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 182063

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 182063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182064)

    End Select
    
    Exit Sub

End Sub

Private Sub Fornecedor_Preenche(objControle As Object)

Static sNomeReduzidoParte As String
Dim lErro As Long
    
On Error GoTo Erro_Fornecedor_Preenche
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objControle, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 182065

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 182065

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182066)

    End Select
    
    Exit Sub

End Sub

Function Limpa_Tela_LivRegES() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_LivRegES
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    OptionCli.Value = True
    
    iClienteAtual = CLIFORN_DE
    iFornecedorAtual = CLIFORN_DE
    iTratamentoAtual = TRATA_CLIENTE
    iAlterado = REGISTRO_INALTERADO

    Limpa_Tela_LivRegES = SUCESSO

    Exit Function

Erro_Limpa_Tela_LivRegES:

    Limpa_Tela_LivRegES = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182067)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal objLivRegESAtualiza As ClassLivRegESAtualiza) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    
    objLivRegESAtualiza.dtData = StrParaDate(Data.Text)
    
    If LivrosFechados.Value = vbChecked Then
        objLivRegESAtualiza.iIncluiLivRegFechados = MARCADO
    Else
        objLivRegESAtualiza.iIncluiLivRegFechados = DESMARCADO
    End If
    
    If OptionCli.Value = True Then
    
        objLivRegESAtualiza.iAtualizaCliente = MARCADO
        objLivRegESAtualiza.lClienteAte = LCodigo_Extrai(CliFornAte.Text)
        objLivRegESAtualiza.lClienteDe = LCodigo_Extrai(CliFornDe.Text)
        
    Else
    
        objLivRegESAtualiza.iAtualizaCliente = DESMARCADO
        objLivRegESAtualiza.lFornecedorAte = LCodigo_Extrai(CliFornAte.Text)
        objLivRegESAtualiza.lFornecedorDe = LCodigo_Extrai(CliFornDe.Text)
        
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182068)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Change(objFornecedorMask As MaskEdBox)
    
    iAlterado = REGISTRO_ALTERADO
    
    Call Fornecedor_Preenche(objFornecedorMask)

End Sub

Private Sub Fornecedor_GotFocus(objFornecedorMask As MaskEdBox)

    Call MaskEdBox_TrataGotFocus(objFornecedorMask, iAlterado)
    
End Sub

Private Sub Fornecedor_Validate(objFornecedorMask As MaskEdBox, Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer

On Error GoTo Erro_Fornecedor_Validate

    If Len(Trim(objFornecedorMask.Text)) > 0 Then

        If LCodigo_Extrai(objFornecedorMask.Text) <> 0 Then
            objFornecedorMask.Text = LCodigo_Extrai(objFornecedorMask.Text)
        End If
        
        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le3(objFornecedorMask, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 182069
        
        objFornecedorMask.Text = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
       
    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 182069
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182070)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornecedor_Click(objFornecedorMask As MaskEdBox)

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedor_Click
    
    If Len(Trim(objFornecedorMask.Text)) > 0 Then
    
        'Preenche com o Fornecedor da tela
        If LCodigo_Extrai(objFornecedorMask.Text) <> 0 Then
        
            objFornecedor.lCodigo = LCodigo_Extrai(objFornecedorMask.Text)
            
        Else
        
            objFornecedor.sNomeReduzido = objFornecedorMask.Text
            
        End If
        
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

   Exit Sub

Erro_LabelFornecedor_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182071)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    If iFornecedorAtual = 1 Then
        CliFornDe.Text = CStr(objFornecedor.lCodigo)
        Call Fornecedor_Validate(CliFornDe, bSGECancelDummy)
    Else
        CliFornAte.Text = CStr(objFornecedor.lCodigo)
        Call Fornecedor_Validate(CliFornAte, bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub Cliente_Change(objClienteMask As MaskEdBox)
    
    iAlterado = REGISTRO_ALTERADO
    
    Call Cliente_Preenche(objClienteMask)

End Sub

Private Sub Cliente_GotFocus(objClienteMask As MaskEdBox)
    
    Call MaskEdBox_TrataGotFocus(objClienteMask, iAlterado)
    
End Sub

Private Sub Cliente_Validate(objClienteMask As MaskEdBox, Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Cliente_Validate

    If Len(Trim(objClienteMask.Text)) > 0 Then
    
        If LCodigo_Extrai(objClienteMask.Text) <> 0 Then
            objClienteMask.Text = LCodigo_Extrai(objClienteMask.Text)
        End If
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(objClienteMask, objCliente, 0)
        If lErro <> SUCESSO Then gError 182072
        
        objClienteMask.Text = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido

    End If
        
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 182072
            'erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182073)

    End Select

End Sub

Private Sub LabelCliente_Click(objClienteMask As MaskEdBox)

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelCliente_Click
        
    'Verifica se Cliente Inicial foi preenchido
    If Len(Trim(objClienteMask.Text)) > 0 Then

        If LCodigo_Extrai(objClienteMask.Text) <> 0 Then

            objCliente.lCodigo = LCodigo_Extrai(objClienteMask.Text)

        Else

            objCliente.sNomeReduzido = objClienteMask.Text

        End If

    End If
        
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

   Exit Sub

Erro_LabelCliente_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182074)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    If iClienteAtual = 1 Then
        CliFornDe.Text = CStr(objCliente.lCodigo)
        Call Cliente_Validate(CliFornDe, bSGECancelDummy)
    Else
        CliFornAte.Text = CStr(objCliente.lCodigo)
        Call Cliente_Validate(CliFornAte, bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub OptionCli_Click()

    If iTratamentoAtual <> TRATA_CLIENTE Then
    
        CliFornDe.Text = ""
        CliFornAte.Text = ""
        
        FrameCliForn.Caption = "Cliente"
    
        iTratamentoAtual = TRATA_CLIENTE
    
    End If

End Sub

Private Sub OptionForn_Click()

    If iTratamentoAtual <> TRATA_FORNECEDOR Then
    
        CliFornDe.Text = ""
        CliFornAte.Text = ""
        
        FrameCliForn.Caption = "Fornecedor"
        
        iTratamentoAtual = TRATA_FORNECEDOR
    
    End If

End Sub

Private Sub LabelCliFornDe_Click()
    
    If iTratamentoAtual = TRATA_CLIENTE Then
        iClienteAtual = CLIFORN_DE
        Call LabelCliente_Click(CliFornDe)
    Else
        iFornecedorAtual = CLIFORN_DE
        Call LabelFornecedor_Click(CliFornDe)
    End If
    
End Sub

Private Sub LabelCliFornAte_Click()
    
    If iTratamentoAtual = TRATA_CLIENTE Then
        iClienteAtual = CLIFORN_ATE
        Call LabelCliente_Click(CliFornAte)
    Else
        iFornecedorAtual = CLIFORN_ATE
        Call LabelFornecedor_Click(CliFornAte)
    End If
    
End Sub

Private Sub CliFornDe_Change()

    If iTratamentoAtual = TRATA_CLIENTE Then
        iClienteAtual = CLIFORN_DE
        Call Cliente_Change(CliFornDe)
    Else
        iFornecedorAtual = CLIFORN_DE
        Call Fornecedor_Change(CliFornDe)
    End If
    
End Sub

Private Sub CliFornAte_Change()

    If iTratamentoAtual = TRATA_CLIENTE Then
        iClienteAtual = CLIFORN_ATE
        Call Cliente_Change(CliFornAte)
    Else
        iFornecedorAtual = CLIFORN_ATE
        Call Fornecedor_Change(CliFornAte)
    End If
    
End Sub

Private Sub CliFornDe_GotFocus()

    If iTratamentoAtual = TRATA_CLIENTE Then
        iClienteAtual = CLIFORN_DE
        Call Cliente_GotFocus(CliFornDe)
    Else
        iFornecedorAtual = CLIFORN_DE
        Call Fornecedor_GotFocus(CliFornDe)
    End If
    
End Sub

Private Sub CliFornAte_GotFocus()

    If iTratamentoAtual = TRATA_CLIENTE Then
        iClienteAtual = CLIFORN_ATE
        Call Cliente_GotFocus(CliFornAte)
    Else
        iFornecedorAtual = CLIFORN_ATE
        Call Fornecedor_GotFocus(CliFornAte)
    End If
    
End Sub

Private Sub CliFornDe_Validate(Cancel As Boolean)

    If iTratamentoAtual = TRATA_CLIENTE Then
        iClienteAtual = CLIFORN_DE
        Call Cliente_Validate(CliFornDe, Cancel)
    Else
        iFornecedorAtual = CLIFORN_DE
        Call Fornecedor_Validate(CliFornDe, Cancel)
    End If
    
End Sub

Private Sub CliFornAte_Validate(Cancel As Boolean)

    If iTratamentoAtual = TRATA_CLIENTE Then
        iClienteAtual = CLIFORN_ATE
        Call Cliente_Validate(CliFornAte, Cancel)
    Else
        iFornecedorAtual = CLIFORN_ATE
        Call Fornecedor_Validate(CliFornAte, Cancel)
    End If
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CliFornDe Then Call LabelCliFornDe_Click
        If Me.ActiveControl Is CliFornAte Then Call LabelCliFornAte_Click
    
    End If
    
End Sub
