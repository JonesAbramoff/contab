VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BorderoCobrancaCancelarOcx 
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   7725
   Begin VB.Frame Frame2 
      Caption         =   "Títulos"
      Height          =   810
      Left            =   135
      TabIndex        =   8
      Top             =   1650
      Width           =   7455
      Begin VB.Label Label2 
         Caption         =   "Quantidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   690
         TabIndex        =   10
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label QtdeTitulos 
         Height          =   225
         Left            =   1830
         TabIndex        =   11
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Label5 
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3900
         TabIndex        =   12
         Top             =   345
         Width           =   525
      End
      Begin VB.Label ValorCobrado 
         Height          =   225
         Left            =   4485
         TabIndex        =   13
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2910
      Picture         =   "BorderoCobrancaCancelarOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   840
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   3945
      Picture         =   "BorderoCobrancaCancelarOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação do Borderô"
      Height          =   1470
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Trazer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6495
         TabIndex        =   4
         Top             =   975
         Width           =   795
      End
      Begin VB.ComboBox Carteira 
         Height          =   315
         Left            =   3945
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2400
      End
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   345
         Width           =   1920
      End
      Begin MSMask.MaskEdBox NumBordero 
         Height          =   330
         Left            =   4920
         TabIndex        =   3
         Top             =   960
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataEmissao 
         Height          =   300
         Left            =   2205
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   945
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   1035
         TabIndex        =   2
         Top             =   960
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label NumBorderoLabel 
         Caption         =   "Número do Borderô:"
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
         Height          =   225
         Left            =   3135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   1005
         Width           =   1740
      End
      Begin VB.Label Label8 
         Caption         =   "Emissão:"
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
         Height          =   255
         Left            =   195
         TabIndex        =   15
         Top             =   975
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Carteira:"
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
         Height          =   255
         Left            =   3165
         TabIndex        =   16
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Cobrador:"
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
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   390
         Width           =   855
      End
   End
End
Attribute VB_Name = "BorderoCobrancaCancelarOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoBorderoCobranca As AdmEvento
Attribute objEventoBorderoCobranca.VB_VarHelpID = -1

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim objBorderoBD As New ClassBorderoCobranca
Dim vbMsgRes  As VbMsgBoxResult

On Error GoTo Erro_BotaoOK_Click
    
    If Len(Trim(Cobrador.Text)) = 0 Then Error 46378
    If Len(Trim(Carteira.Text)) = 0 Then Error 46379
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Error 46380
    If Len(Trim(NumBordero.Text)) = 0 Then Error 46381
    
    objBorderoCobranca.iCobrador = Codigo_Extrai(Cobrador.Text)
    objBorderoCobranca.iCodCarteiraCobranca = Codigo_Extrai(Carteira.Text)
    objBorderoCobranca.dtDataEmissao = MaskedParaDate(DataEmissao)
    objBorderoCobranca.lNumBordero = CLng(NumBordero.Text)
    
    objBorderoBD.lNumBordero = objBorderoCobranca.lNumBordero
    
    lErro = CF("BorderoCobranca_Le",objBorderoBD)
    If lErro <> SUCESSO And lErro <> 46366 Then Error 46382
    If lErro <> SUCESSO Then Error 46383
    
    If objBorderoBD.iCobrador <> objBorderoCobranca.iCobrador Or objBorderoBD.iCodCarteiraCobranca <> objBorderoCobranca.iCodCarteiraCobranca _
       Or objBorderoBD.dtDataEmissao <> objBorderoCobranca.dtDataEmissao Then
       
       vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BORDERO_BD_DIFERENTE", objBorderoCobranca.lNumBordero, objBorderoBD.iCobrador, objBorderoBD.iCodCarteiraCobranca, objBorderoBD.dtDataEmissao)
       If vbMsgRes = vbNo Then Error 46384
            
        lErro = Carrega_Dados_BorderoCobranca()
        If lErro <> SUCESSO Then Error 46413
        
    End If
    
    If objBorderoBD.iStatus = STATUS_EXCLUIDO Then Error 46386
    
    objBorderoCobranca.dtDataCancelamento = gdtDataHoje
    objBorderoCobranca.dtDataContabilCancelamento = gdtDataAtual
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("BorderoCobranca_Cancelar",objBorderoCobranca)
    If lErro <> SUCESSO Then Error 46377
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Call Limpa_Tela_BorderoCobrancaCancelar
    
    Exit Sub
    
Erro_BotaoOK_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 46377, 46382, 46413
        
        Case 46378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)
        
        Case 46379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBRANCA_NAO_INFORMADA", Err)
        
        Case 46380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
            
        Case 46381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_NAO_INFORMADO", Err)
            
        Case 46383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERO_COBRANCA_NAO_CADASTRADO", Err, objBorderoBD.lNumBordero)
            
        Case 46386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERO_COBRANCA_EXCLUIDO", Err, objBorderoBD.lNumBordero)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143675)
            
    End Select
    
    Exit Sub
        
End Sub

Private Sub BotaoTrazer_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoTrazer_Click

    If Len(Trim(Cobrador.Text)) = 0 Or Len(Trim(Carteira.Text)) = 0 _
        Or Len(Trim(DataEmissao.ClipText)) = 0 Or Len(Trim(NumBordero.ClipText)) = 0 Then Error 56753
    
    lErro = Carrega_Dados_BorderoCobranca()
    If lErro <> SUCESSO Then Error 46362
    
    Exit Sub
    
Erro_BotaoTrazer_Click:

    Select Case Err
    
        Case 56753
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCH_CPOS_OBRIG_TELA", Err)
        
        Case 46362
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143676)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Dados_Bordero()

    QtdeTitulos.Caption = ""
    ValorCobrado.Caption = ""

End Sub

Private Sub Carteira_Change()
    
    Call Limpa_Dados_Bordero
    
End Sub

Private Sub Cobrador_Change()

    Call Limpa_Dados_Bordero

End Sub

Private Sub DataEmissao_Change()

    Call Limpa_Dados_Bordero

End Sub

Private Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 46353

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 46353

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143677)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoBorderoCobranca = New AdmEvento

    lErro = Carrega_Cobradores()
    If lErro <> SUCESSO Then Error 46351
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    Select Case Err
    
        Case 46351
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143678)
    
    End Select
    
    Exit Sub

End Sub

Private Function Carrega_Cobradores() As Long

Dim lErro As Long
Dim objCobrador As ClassCobrador
Dim ColCobrador As New Collection

On Error GoTo Erro_Carrega_Cobradores

    'Carrega a Coleção de Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial",ColCobrador)
    If lErro <> SUCESSO Then Error 46352
    
    'Preenche a ComboBox Cobrador com os objetos da coleção de Cobradores
    For Each objCobrador In ColCobrador

        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        End If

    Next

    Carrega_Cobradores = SUCESSO
    
    Exit Function
    
Erro_Carrega_Cobradores:

    Carrega_Cobradores = Err
    
    Select Case Err
    
        Case 46352
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143679)
            
    End Select
    
    Exit Function

End Function

Private Sub Cobrador_Click()

Dim iCodCobrador As Integer
Dim objCobrador As New ClassCobrador
Dim lErro As Long
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim sListBoxItem As String
Dim colCarteirasCobrador As New Collection

On Error GoTo Erro_Cobrador_Click
    
    'Limpa a Combo de Carteiras
    Carteira.Clear

    If Cobrador.ListIndex = -1 Then Exit Sub
    
    'Se Cobrador está preenchido
    If Len(Trim(Cobrador.Text)) <> 0 Then

        'Extrai o código do Cobrador
        iCodCobrador = Codigo_Extrai(Cobrador.Text)
    
        'Passa o Código do Cobrador que está na tela para o Obj
        objCobrador.iCodigo = iCodCobrador
    
        'Lê os dados do Cobrador
        lErro = CF("Cobrador_Le",objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 46354
    
        'Se o Cobrador não estiver cadastrado
        If lErro = 19294 Then Error 46358
                                
        'Le as carteiras associadas ao Cobrador
        lErro = CF("Cobrador_Le_Carteiras",objCobrador, colCarteirasCobrador)
        If lErro <> SUCESSO And lErro <> 23500 Then Error 46356

        If lErro = SUCESSO Then
        
            'Preencher a Combo
            For Each objCarteiraCobrador In colCarteirasCobrador
                       
                objCarteiraCobranca.iCodigo = objCarteiraCobrador.iCodCarteiraCobranca
        
                lErro = CF("CarteiraDeCobranca_Le",objCarteiraCobranca)
                If lErro <> SUCESSO And lErro <> 23413 Then Error 46357
        
                'Carteira não está cadastrado
                If lErro = 23413 Then Error 46359
       
                'Concatena Código e a Descricao da carteira
                sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
                sListBoxItem = sListBoxItem & SEPARADOR & objCarteiraCobranca.sDescricao
              
                Carteira.AddItem sListBoxItem
                Carteira.ItemData(Carteira.NewIndex) = objCarteiraCobranca.iCodigo
            Next
        End If
                
        If Carteira.ListCount <> 0 Then Carteira.ListIndex = 0

    End If

    Exit Sub

Erro_Cobrador_Click:

    Select Case Err

        Case 46354, 46356, 46357
            Cobrador.SetFocus
        
        Case 46358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", Err, Cobrador.Text)
            Cobrador.SetFocus
              
        Case 46359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err, objCarteiraCobranca.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143680)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoBorderoCobranca = Nothing

End Sub

Private Sub NumBordero_Change()

    Call Limpa_Dados_Bordero

End Sub

Private Sub NumBordero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumBordero)

End Sub

Private Sub NumBordero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumBordero_Validate

    If Len(Trim(NumBordero.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(NumBordero.Text)
        If lErro <> SUCESSO Then Error 57999
        
    End If
    
    Exit Sub
    
Erro_NumBordero_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 57999 'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143681)
    
    End Select
    
    Exit Sub

End Sub

Private Sub NumBorderoLabel_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objBorderoCobranca As New ClassBorderoCobranca

On Error GoTo Erro_NumBorderoLabel_Click

    If Len(Trim(Cobrador.Text)) = 0 Then Error 46408
    If Len(Trim(Carteira.Text)) = 0 Then Error 46409
    
    colSelecao.Add Codigo_Extrai(Cobrador.Text)
    colSelecao.Add Codigo_Extrai(Carteira.Text)
    
    If Len(Trim(DataEmissao.ClipText)) <> 0 Then objBorderoCobranca.dtDataEmissao = MaskedParaDate(DataEmissao)
    If Len(Trim(NumBordero.Text)) <> 0 Then objBorderoCobranca.lNumBordero = CLng(NumBordero.Text)
    
    Call Chama_Tela("BorderoCobrancaLista", colSelecao, objBorderoCobranca, objEventoBorderoCobranca)

    Exit Sub

Erro_NumBorderoLabel_Click:

    Select Case Err
    
        Case 46408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)
            
        Case 46409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBRANCA_NAO_INFORMADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143682)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoBorderoCobranca_evSelecao(obj1 As Object)
    
Dim objBorderoCobranca As ClassBorderoCobranca
Dim iIndice As Integer

    Set objBorderoCobranca = obj1
    
    Call Limpa_Tela_BorderoCobrancaCancelar
    
    NumBordero.Text = objBorderoCobranca.lNumBordero
    
    Call Carrega_Dados_BorderoCobranca

    Me.Show
    
End Sub

Private Sub UpDownDataEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_DownClick

    'Diminui a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 46360

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case Err

        Case 46360

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143683)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_UpClick

    'Aumenta a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 46361

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case Err

        Case 46361

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143684)

    End Select

    Exit Sub

End Sub

Function Carrega_Dados_BorderoCobranca() As Long
'traz para a tela dados do bordero de cobranca identificado pelo numero na tela

Dim lErro As Long
Dim objBorderoCobranca As New ClassBorderoCobranca
Dim lQuantidade As Long
Dim dValorCobrado As Double
Dim dtDataEmissao As Date
Dim iCobrador As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Dados_BorderoCobranca

    dtDataEmissao = MaskedParaDate(DataEmissao)
    iCobrador = Codigo_Extrai(Cobrador)
    objBorderoCobranca.lNumBordero = CLng(NumBordero.Text)
    
    lErro = CF("BorderoCobranca_Le",objBorderoCobranca)
    If lErro <> SUCESSO And lErro <> 46366 Then Error 46363
    If lErro <> SUCESSO Then Error 46367
    
    lErro = CF("BorderoCobranca_Le_Quantidade_ValorCobrado",objBorderoCobranca, lQuantidade, dValorCobrado)
    If lErro <> SUCESSO Then Error 46368
       
    For iIndice = 0 To Cobrador.ListCount
        If Cobrador.ItemData(iIndice) = objBorderoCobranca.iCobrador Then
            Cobrador.ListIndex = iIndice
            Exit For
        End If
    Next
    
    For iIndice = 0 To Carteira.ListCount
        If Carteira.ItemData(iIndice) = objBorderoCobranca.iCodCarteiraCobranca Then
            Carteira.ListIndex = iIndice
            Exit For
        End If
    Next
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(objBorderoCobranca.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    QtdeTitulos.Caption = lQuantidade
    ValorCobrado.Caption = Format(dValorCobrado, "Standard")

    Carrega_Dados_BorderoCobranca = SUCESSO

    Exit Function
    
Erro_Carrega_Dados_BorderoCobranca:

    Carrega_Dados_BorderoCobranca = Err
    
    Select Case Err
    
        Case 46363, 46368
        
        Case 46367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERO_COBRANCA_NAO_CADASTRADO", Err, objBorderoCobranca.lNumBordero)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143685)
            
    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_BorderoCobrancaCancelar()

    Call Limpa_Tela(Me)
    
    Cobrador.ListIndex = -1
    QtdeTitulos.Caption = ""
    ValorCobrado.Caption = ""

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CANCELAR_BORDERO_COBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Cancelar Borderô de Cobrança"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoCobrancaCancelar"
    
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

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NumBordero Then
            Call NumBorderoLabel_Click
        End If
    
    End If
    
End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub QtdeTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdeTitulos, Source, X, Y)
End Sub

Private Sub QtdeTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdeTitulos, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub ValorCobrado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorCobrado, Source, X, Y)
End Sub

Private Sub ValorCobrado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorCobrado, Button, Shift, X, Y)
End Sub

Private Sub NumBorderoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumBorderoLabel, Source, X, Y)
End Sub

Private Sub NumBorderoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumBorderoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

