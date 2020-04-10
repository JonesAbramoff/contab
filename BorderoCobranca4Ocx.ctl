VERSION 5.00
Begin VB.UserControl BorderoCobranca4Ocx 
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   ScaleHeight     =   3975
   ScaleWidth      =   5910
   Begin VB.Frame Frame2 
      Caption         =   "Bordero Gerado"
      Height          =   750
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   5760
      Begin VB.Label labelBordero 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3120
         TabIndex        =   10
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Número do Borderô"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1335
         TabIndex        =   9
         Top             =   330
         Width           =   1665
      End
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   2325
      Picture         =   "BorderoCobranca4Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3270
      Width           =   960
   End
   Begin VB.Frame Impressao 
      Caption         =   "Saídas em"
      Height          =   2355
      Left            =   90
      TabIndex        =   5
      Top             =   870
      Width           =   5760
      Begin VB.Frame Frame1 
         Caption         =   "Localização do Arquivo"
         Height          =   1455
         Left            =   105
         TabIndex        =   6
         Top             =   750
         Width           =   5550
         Begin VB.CommandButton BotaoProcurar 
            Caption         =   "..."
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
            Left            =   4890
            TabIndex        =   3
            Top             =   960
            Width           =   555
         End
         Begin VB.CheckBox ArquivoCNAB 
            Caption         =   "Arquivo CNAB"
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
            Left            =   135
            TabIndex        =   1
            Top             =   375
            Width           =   1575
         End
         Begin VB.TextBox NomeArquivo 
            Height          =   315
            Left            =   75
            TabIndex        =   2
            Top             =   1005
            Width           =   4800
         End
         Begin VB.Label Label1 
            Caption         =   "Localização do Arquivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   750
            Width           =   2145
         End
      End
      Begin VB.CheckBox CheckImpressora 
         Caption         =   "Impressora"
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
         Left            =   255
         TabIndex        =   0
         Top             =   315
         Width           =   1335
      End
   End
End
Attribute VB_Name = "BorderoCobranca4Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Declare Function CNAB_PagRem_Abrir Lib "ADCNAB.DLL" (lCNABPagRem As Long, ByVal sNomeArq As String, ByVal iCodigoBanco As Integer, ByVal iNumRemessa As Integer, vDataEmissao As Variant, ByVal iTipoCobranca As Integer, ByVal iLiqTitOutrosBcos As Integer) As Long
Private Declare Function CNAB_PagRem_Fechar Lib "ADCNAB.DLL" (ByVal lCNABPagRem As Long) As Long
Private Declare Function CNAB_PagRem_DefCtaEmpresa Lib "ADCNAB.DLL" (ByVal lCNABPagRem As Long, ByVal sAgencia As String, ByVal sConta As String, ByVal sDVConta As String) As Long
Private Declare Function CNAB_PagRem_IncluirReg Lib "ADCNAB.DLL" (ByVal lCNABPagRem As Long, ByVal dValorBaixado As Double, ByVal sSiglaDocumento As String, vDataVencimento As Variant, vDataEmissao As Variant, ByVal sPagtoId As String, ByVal lNumTitulo As Long, ByVal sNossoNumero As String, _
ByVal sEndereco As String, ByVal sCidade As String, ByVal sSiglaEstado As String, ByVal sCEP As String, ByVal sCgc As String, ByVal sRazaoSocial As String, ByVal iBanco As Integer, ByVal sAgencia As String, ByVal sContaCorrente As String) As Integer

Dim iListIndexDefault As Integer
Dim gobjBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao

Private Sub ArquivoCNAB_Click()

    If ArquivoCNAB.Value = vbChecked Then
        'Dir1.Enabled = True
        'Drive1.Enabled = True
        NomeArquivo.Enabled = True
        BotaoProcurar.Enabled = True
    Else
        'Dir1.Enabled = False
        'Drive1.Enabled = False
        BotaoProcurar.Enabled = False
        NomeArquivo.Enabled = False
    End If

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim sNomeArqParam As String
Dim sNomeDir As String
Dim objCobrancaEletronica As New ClassCobrancaEletronica

On Error GoTo Erro_BotaoOK_Click

    If ArquivoCNAB.Value = vbChecked Then
        If Len(Trim(NomeArquivo.Text)) = 0 Then gError 80343
    End If

    If CheckImpressora.Value = vbChecked Then
        lErro = ImprimirBordero(gobjBorderoCobrancaEmissao.lNumero)
        If lErro <> SUCESSO Then gError 80344
    End If

'''    'se for p/criar o arq cnab...
'''    If ArquivoCNAB.Value = vbChecked Then
'''
'''        sNomeDir = NomeArquivo.Text
'''
'''        lErro = Sistema_Preparar_Batch(sNomeArqParam)
'''        If lErro <> SUCESSO Then gError 80345
'''
'''        lErro = CF("BorderoPagto_Abre_TelaRemessaArq",sNomeArqParam, sNomeDir, gobjBorderoCobrancaEmissao)
'''        If lErro <> SUCESSO Then gError 80346
'''
'''    End If

'Maristela(Inicio)
'''    'se for p/criar o arq cnab...
'''    If ArquivoCNAB.Value = vbChecked Then
'''
'''        sDiretorio = NomeArquivo.Text
'''
'''        lErro = Sistema_Preparar_Batch(sNomeArqParam)
'''        If lErro <> SUCESSO Then gError 80345
'''
'''        lErro = CobrancaEletronica_Le_TelaGeracaoArq(sNomeArqParam, sDiretorio, gobjBorderoCobrancaEmissao)
'''        If lErro <> SUCESSO Then gError 80346
'''
'''    End If
'Maristela(Fim)

    'se for p/criar o arq cnab...
    If ArquivoCNAB.Value = vbChecked Then
    
        'Carrega o objCobrancaEletronica com o cobrador
        objCobrancaEletronica.iCobrador = gobjBorderoCobrancaEmissao.iCobrador
        objCobrancaEletronica.objCobrador.iCodigo = objCobrancaEletronica.iCobrador
    
        'Lê os dados do cobrador
        lErro = CF("Cobrador_Le", objCobrancaEletronica.objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 51642
        If lErro <> SUCESSO Then gError 51643
        
        objCobrancaEletronica.iNumBorderoIni = gobjBorderoCobrancaEmissao.lNumero
        objCobrancaEletronica.iNumBorderoFim = gobjBorderoCobrancaEmissao.lNumero
        objCobrancaEletronica.iRegerarArquivo = vbUnchecked
        
        'Lê os registros em OcorrRemParcRec
        lErro = CF("CobrancaEletronica_Obter_Borderos", objCobrancaEletronica)
        If lErro <> SUCESSO Then gError 51644
           
        'Se não encontrou --> erro
        If objCobrancaEletronica.colBorderos.Count = 0 Then gError 51645
        
        objCobrancaEletronica.sDiretorio = Trim(NomeArquivo.Text)
    
        lErro = CF("CobrancaEletronica_Abre_TelaGeracaoArq", sNomeArqParam, objCobrancaEletronica)
        If lErro <> SUCESSO Then gError 62286

    End If

    'Fecha a tela
    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case 80344, 80345, 80346
        
        Case 80343
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
            NomeArquivo.SetFocus

        Case 51641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
                
        Case 51642, 51644, 62286, 62287
        
        Case 51643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, objCobrancaEletronica.objCobrador)
        
        Case 51645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_BORDEROSCOBRANCA", gErr)
        
        Case 93564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BORDERODE_MAIOR_BORDEROATE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143668)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long
'Traz os dados das Parcelas a pagar para a Tela

Dim sDiretorio As String

    Set gobjBorderoCobrancaEmissao = objBorderoCobrancaEmissao

    labelBordero.Caption = objBorderoCobrancaEmissao.lNumero

    Call CF("BancosInfo_Diretorio_Le", sDiretorio, gobjBorderoCobrancaEmissao.iCobrador)

    NomeArquivo.Text = sDiretorio
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Function ImprimirBordero(lNumero As Long) As Long
'chama a impressao de bordero

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String, sBuffer As String
Dim lErro As Long

On Error GoTo Erro_ImprimirBordero

    lErro = objRelatorio.ExecutarDireto("Borderô de Cobrança", "", 0, "", "NBORDERO", CStr(lNumero), "NCOBRADOR", gobjBorderoCobrancaEmissao.iCobrador)
    If lErro <> SUCESSO Then gError 80347

    ImprimirBordero = SUCESSO

    Exit Function

Erro_ImprimirBordero:

    ImprimirBordero = gErr

    Select Case gErr

        Case 80347

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143669)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim sDiretorio As String

On Error GoTo Erro_Form_Load

    'NomeArquivo.Text = Dir1.Path
    'iListIndexDefault = Drive1.ListIndex
    
'    lErro = CF("BancosInfo_Diretorio_Le", sDiretorio, gobjBorderoCobrancaEmissao.iCobrador)
'    If lErro <> SUCESSO Then gError 99999
'
'    NomeArquivo.Text = sDiretorio

    BotaoProcurar.Enabled = False
    NomeArquivo.Enabled = False

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        '######################################
        'Inserido por Wagner - REPLICAR_ACERTO
        Case 68, 76
        
            sDiretorio = CurDir
            'Dir1.Path = sDiretorio
            NomeArquivo.Text = sDiretorio
            Resume Next
        '######################################
    
        Case 99999

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143670)

    End Select

    Exit Sub

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set gobjBorderoCobrancaEmissao = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_PAGT_P4
    Set Form_Load_Ocx = Me
    Caption = "Bordero de Cobranças - Saidas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoCobranca4"

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

Private Sub NomeArquivo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeArquivo_Validate

    If Len(Trim(NomeArquivo.Text)) = 0 Then Exit Sub

    If Len(Trim(Dir(NomeArquivo.Text, vbDirectory))) = 0 Then gError 80348

    'Drive1.Drive = Mid(NomeArquivo, 1, 2)

    'Dir1.Path = NomeArquivo.Text

    Exit Sub

Erro_NomeArquivo_Validate:

    Cancel = True


    Select Case gErr

        Case 80348
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeArquivo.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143671)

    End Select

    Exit Sub

End Sub

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

'
'Private Sub Drive1_Change()
'
'On Error GoTo Erro_Drive1_Change
'
'    Dir1.Path = Drive1.Drive
'
'    Exit Sub
'
'Erro_Drive1_Change:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143672)
'
'    End Select
'
'    Drive1.ListIndex = iListIndexDefault
'
'    Exit Sub
'
'End Sub
'
'Private Sub Drive1_GotFocus()
'
'    iListIndexDefault = Drive1.ListIndex
'
'End Sub
'
'Private Sub Dir1_Change()
'
'    NomeArquivo = Dir1.Path
'
'End Sub
'
'
'Private Sub Dir1_Click()
'
'On Error GoTo Erro_Dir1_Click
'
'    Exit Sub
'
'Erro_Dir1_Click:
'
'    Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143673)
'
'    Exit Sub
'
'End Sub


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'Maristela(Inicio)
'''Public Function CobrancaEletronica_Le_TelaGeracaoArq(sNomeArqParam As String, sDiretorio As String, objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long
''''Prepara o ObjCobrancaEletronica
'''
'''Dim lErro As Long
'''Dim objCobrancaEletronica As New ClassCobrancaEletronica
'''
'''
'''On Error GoTo Erro_CobrancaEletronica_Le_TelaGeracaoArq
'''
'''    'Carrega o objCobrancaEletronica com o cobrador
'''    objCobrancaEletronica.objCobrador.iCodigo = objBorderoCobrancaEmissao.iCobrador
'''
'''    'Lê os dados do cobrador
'''    lErro = CF("Cobrador_Le",objCobrancaEletronica.objCobrador)
'''    If lErro <> SUCESSO And lErro <> 19294 Then Error 51642
'''    If lErro <> SUCESSO Then Error 51643
'''
'''    'Lê os registros em OcorrRemParcRec
'''    objCobrancaEletronica.iCobrador = objBorderoCobrancaEmissao.iCobrador
'''    objCobrancaEletronica.sDiretorio = sDiretorio
'''
'''    lErro = CF("CobrancaEletronica_Abre_TelaGeracaoArq",sNomeArqParam, objCobrancaEletronica)
'''    If lErro <> SUCESSO Then Error 62286
'''
'''    Unload Me
'''
'''    Exit Function
'''
'''
'''Erro_CobrancaEletronica_Le_TelaGeracaoArq:
'''
'''    Select Case Err
'''
'''        Case 51641
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)
'''
'''        Case 51642, 51644, 62286, 62287
'''
'''        Case 51643
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, objCobrancaEletronica.objCobrador)
'''
'''        Case 51645
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_BORDEROSCOBRANCA", Err)
'''
'''        Case Else
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143674)
'''
'''    End Select
'''
'''    Exit Function
'''
'''End Function
'Maristela(Fim)

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização dos arquivos do borderô"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = String(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeArquivo.Text = sBuffer
        Call NomeArquivo_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub
