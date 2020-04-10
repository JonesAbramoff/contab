VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl GeracaoArqRemCobrOcx 
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   ScaleHeight     =   3645
   ScaleWidth      =   4860
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
      Left            =   4050
      TabIndex        =   4
      Top             =   1980
      Width           =   555
   End
   Begin VB.CheckBox Regerar 
      Caption         =   "Gerando Novamente o Arquivo"
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
      Left            =   255
      TabIndex        =   12
      Top             =   2445
      Width           =   3180
   End
   Begin VB.Frame Frame1 
      Caption         =   "Borderô"
      Height          =   705
      Left            =   195
      TabIndex        =   7
      Top             =   825
      Width           =   4455
      Begin MSMask.MaskEdBox BorderoDe 
         Height          =   300
         Left            =   1005
         TabIndex        =   10
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox BorderoAte 
         Height          =   300
         Left            =   2820
         TabIndex        =   11
         Top             =   255
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelBorderoAte 
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
         Height          =   195
         Left            =   2355
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   315
         Width           =   360
      End
      Begin VB.Label LabelBorderoDe 
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
         Height          =   195
         Left            =   630
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.TextBox NomeDiretorio 
      Height          =   330
      Left            =   255
      TabIndex        =   3
      Top             =   2010
      Width           =   3810
   End
   Begin VB.ComboBox Cobrador 
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   3450
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   2550
      Picture         =   "GeracaoArqRemCobrOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2895
      Width           =   855
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   1620
      Picture         =   "GeracaoArqRemCobrOcx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Localização do Arquivo"
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
      Height          =   285
      Left            =   270
      TabIndex        =   5
      Top             =   1725
      Width           =   2145
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   330
      TabIndex        =   6
      Top             =   300
      Width           =   840
   End
End
Attribute VB_Name = "GeracaoArqRemCobrOcx"
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
Public iAlterado As Integer
Dim iListIndexDefault As Integer
Dim m_Caption As String

Private WithEvents objEventoBorderoAte As AdmEvento
Attribute objEventoBorderoAte.VB_VarHelpID = -1
Private WithEvents objEventoBorderoDe As AdmEvento
Attribute objEventoBorderoDe.VB_VarHelpID = -1

Event Unload()

Private Sub BorderoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(BorderoAte)

End Sub

Private Sub BorderoDe_GotFocus()
        
    Call MaskEdBox_TrataGotFocus(BorderoDe)
        
End Sub

Private Sub BotaoCancela_Click()
    
    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objCobrancaEletronica As New ClassCobrancaEletronica
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoOK_Click
    
    'validacao dos nº de borderos
    If Len(Trim(BorderoDe.Text)) > 0 And Len(Trim(BorderoAte.Text)) > 0 Then
        If StrParaInt(Trim(BorderoAte.Text)) < StrParaInt(Trim(BorderoDe.Text)) Then gError 93564
    End If
    
    'Verifica se o Cobrador foi selecionado
    If Cobrador.ListIndex = -1 Then gError 51641
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 7784
    
    'Carrega o objCobrancaEletronica com o cobrador
    objCobrancaEletronica.iCobrador = Codigo_Extrai(Cobrador)
    objCobrancaEletronica.objCobrador.iCodigo = objCobrancaEletronica.iCobrador

    'Lê os dados do cobrador
    lErro = CF("Cobrador_Le", objCobrancaEletronica.objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then gError 51642
    If lErro <> SUCESSO Then gError 51643
    
    If Len(Trim(BorderoDe.ClipText)) <> 0 Then objCobrancaEletronica.iNumBorderoIni = StrParaInt(BorderoDe.Text)
    If Len(Trim(BorderoAte.ClipText)) <> 0 Then objCobrancaEletronica.iNumBorderoFim = StrParaInt(BorderoAte.Text)
    objCobrancaEletronica.iRegerarArquivo = Regerar.Value
    
    'Lê os registros em OcorrRemParcRec
    lErro = CF("CobrancaEletronica_Obter_Borderos", objCobrancaEletronica)
    If lErro <> SUCESSO Then gError 51644
       
    'Se não encontrou --> erro
    If objCobrancaEletronica.colBorderos.Count = 0 Then gError 51645
    
    objCobrancaEletronica.sDiretorio = Trim(NomeDiretorio.Text)
    
''    lErro = Sistema_Preparar_Batch(sNomeArqParam)
''    If lErro <> SUCESSO Then gError 62287
    
    lErro = CF("CobrancaEletronica_Abre_TelaGeracaoArq", sNomeArqParam, objCobrancaEletronica)
    If lErro <> SUCESSO Then gError 62286
        
    Unload Me
       
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
    
        Case 7784
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
            NomeDiretorio.SetFocus
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160777)
            
    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim ColCobrador As New Collection
Dim objCobrador As ClassCobrador
Dim lErro As Long
Dim sDiretorio As String

On Error GoTo Erro_Form_Load
    
    Set objEventoBorderoAte = New AdmEvento
    Set objEventoBorderoDe = New AdmEvento
    
'    iListIndexDefault = Drive1.ListIndex
    
'    lErro = CF("BancosInfo_Diretorio_Le", sDiretorio)
'    If lErro <> SUCESSO Then gError 99999
    
'    If Len(Trim(sDiretorio)) > 0 Then
'        Dir1.Path = sDiretorio
'        Drive1.Drive = Left(sDiretorio, 2)
'    End If
'
'    NomeDiretorio = Dir1.Path
    
'    NomeDiretorio = sDiretorio 'Dir1.Path
    
    'Carrega a Coleção de Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then gError 51649
    
    For Each objCobrador In ColCobrador
        
        'Seleciona os cobradores ativos que utilizem cobrança eletrônica
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iInativo <> Inativo And objCobrador.iCobrancaEletronica = vbChecked Then
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        End If

    Next
    
    lErro = CF("BancosInfo_Diretorio_Le", sDiretorio, Codigo_Extrai(Cobrador.Text))
    If lErro <> SUCESSO Then gError 99999
    
    NomeDiretorio.Text = sDiretorio
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 51649, 99999
        
        '######################################
        'Inserido por Wagner - REPLICAR_ACERTO
        Case 68, 76
        
            sDiretorio = CurDir
            'Dir1.Path = sDiretorio
            NomeDiretorio.Text = sDiretorio
            Resume Next
        '######################################
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160778)
    
    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_GERACAO_ARQREMCOBR
    Set Form_Load_Ocx = Me
    Caption = "Seleção do Cobrador - CNAB"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GeracaoArqRemCobr"
    
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

Private Sub Cobrador_Click()

Dim objCobrador As New ClassCobrador
Dim lErro As Long
Dim sListBoxItem As String
Dim lNumMaxBordero As Long
Dim lNumMinBordero As Long
Dim sDiretorio As String

On Error GoTo Erro_Cobrador_Click
    
    'Se Cobrador está preenchido
    If Cobrador.ListIndex = -1 Then Exit Sub
    
    'Passa o Código do Cobrador que está na tela para o Obj
    objCobrador.iCodigo = Cobrador.ItemData(Cobrador.ListIndex)

    'Lê os dados do Cobrador
    lErro = Cobrador_Le_BorderosArqRem(objCobrador, lNumMinBordero, lNumMaxBordero)
    If lErro <> SUCESSO Then gError 93565

    If lNumMinBordero <> 0 Then
        BorderoDe.Text = CStr(lNumMinBordero)
        BorderoAte.Text = CStr(lNumMaxBordero)
    Else
        BorderoDe.Text = ""
        BorderoAte.Text = ""
    End If
    
    lErro = CF("BancosInfo_Diretorio_Le", sDiretorio, Codigo_Extrai(Cobrador.Text))
    If lErro <> SUCESSO Then gError 99999
    
    NomeDiretorio.Text = sDiretorio
            
    Exit Sub

Erro_Cobrador_Click:

    Select Case gErr

        Case 93565
              
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160779)

    End Select

    Exit Sub
    
End Sub
'
'Private Sub Dir1_Change()
'
'    NomeDiretorio = Dir1.Path
'
'End Sub
'
'Private Sub Dir1_Click()
'
'On Error GoTo Erro_Dir1_Click
'
'    Exit Sub
'
'Erro_Dir1_Click:
'
'    Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160780)
'
'    Exit Sub
'
'End Sub
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
'    Select Case Err
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160781)
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

Private Sub LabelBorderoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objBorderoCobranca As New ClassBorderoCobranca

On Error GoTo Erro_LabelBorderoDe_Click
              
    'Verifica se o cobrador foi preenchido
    If Len(Trim(Cobrador.Text)) = 0 Then gError 93572
                 
    If Len(Trim(BorderoDe.Text)) > 0 Then objBorderoCobranca.lNumBordero = CLng(BorderoDe.Text)
    
    colSelecao.Add Codigo_Extrai(Cobrador.Text)
    
    'Chama Tela BorderoCobrancaLista
    Call Chama_Tela("BorderoDeCobrancaLista", colSelecao, objBorderoCobranca, objEventoBorderoDe)
        
    Exit Sub
    
Erro_LabelBorderoDe_Click:

    Select Case gErr

        Case 93572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr, Error$)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160782)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelBorderoDe_DragDrop(Source As Control, X As Single, Y As Single)
     
    Call Controle_DragDrop(LabelBorderoDe, Source, X, Y)
    
End Sub

Private Sub LabelBorderoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
       
       Call Controle_MouseDown(LabelBorderoDe, Button, Shift, X, Y)

End Sub

Private Sub LabelBorderoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objBorderoCobranca As New ClassBorderoCobranca

On Error GoTo Erro_LabelBorderoAte_Click
              
    'Verifica se o cobrador foi preenchido
    If Len(Trim(Cobrador.Text)) = 0 Then gError 93573
              
    If Len(Trim(BorderoAte.Text)) > 0 Then objBorderoCobranca.lNumBordero = CLng(BorderoAte.Text)
    
    colSelecao.Add Codigo_Extrai(Cobrador.Text)
    
    'Chama Tela BorderoCobrancaLista
    Call Chama_Tela("BorderoDeCobrancaLista", colSelecao, objBorderoCobranca, objEventoBorderoAte)

Exit Sub

Erro_LabelBorderoAte_Click:

    Select Case gErr

        Case 93573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr, Error$)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160783)

    End Select

    Exit Sub


End Sub

Private Sub LabelBorderoAte_DragDrop(Source As Control, X As Single, Y As Single)
       
       Call Controle_DragDrop(BorderoAte, Source, X, Y)

End Sub

Private Sub LabelBorderoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
       
       Call Controle_MouseDown(BorderoAte, Button, Shift, X, Y)

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
    
   RaiseEvent Unload
   
   Set objEventoBorderoAte = Nothing
   Set objEventoBorderoDe = Nothing
   
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******

Public Function Trata_Parametros(Optional objCobrador As ClassCobrador) As Long
    
    Trata_Parametros = SUCESSO
    iAlterado = 0
    
    Exit Function

End Function

Private Sub Calcula_DAC(sCampo As String, sDac As String)

Dim iIndice As Integer
Dim iTamCampo As Integer
Dim iDigito1 As Integer
Dim iVarModulo As Integer
Dim iProduto As Integer
Dim iSoma As Integer
Dim sSoma As String
Dim iResto As Integer
Dim iResultado As Integer
    
    iVarModulo = 2
    iSoma = 0
    sSoma = ""
    
    iTamCampo = Len(sCampo)
    
    For iIndice = iTamCampo To 1 Step -1

        iDigito1 = StrParaInt(Mid(sCampo, iIndice, 1))
        iProduto = (iDigito1 * iVarModulo)
        sSoma = sSoma & iProduto
                
        iVarModulo = iVarModulo - 1
        
        If iVarModulo = 0 Then iVarModulo = 2
    
    Next
    
    For iIndice = 1 To Len(sSoma)
        iSoma = iSoma + StrParaInt(Mid(sSoma, iIndice, 1))
    Next
    
    iResto = iSoma Mod 10
    
    iResultado = 10 - iResto
    
    If iResto = 0 Then iResultado = 0
    
    sDac = iResultado
    
    Exit Sub
    
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

Private Sub objEventoBorderoAte_evSelecao(obj1 As Object)

Dim objBorderoCobranca As ClassBorderoCobranca

    Set objBorderoCobranca = obj1

    BorderoAte.Text = CStr(objBorderoCobranca.lNumBordero)

    Me.Show

End Sub

Private Sub objEventoBorderoDe_evSelecao(obj1 As Object)

Dim objBorderoCobranca As ClassBorderoCobranca

    Set objBorderoCobranca = obj1

    BorderoDe.Text = CStr(objBorderoCobranca.lNumBordero)

    Me.Show

End Sub

'??? mover p/cprselect
Function Cobrador_Le_BorderosArqRem(objCobrador As ClassCobrador, lNumMinBordero As Long, lNumMaxBordero As Long) As Long
'Obtem os numeros do menor e do maior bordero do cobrador que ainda nao foram processados (ainda nao foi gerado o arquivo remessa)

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_Cobrador_Le_BorderosArqRem

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 93567

    'Verifica se existe o numero de Bordero para um determinado cobrador
    lErro = Comando_Executar(lComando, "SELECT  MIN(NumBordero), Max(NumBordero) FROM BorderosCobranca WHERE Processado = 0 AND Cobrador = ? AND Status <> ?", lNumMinBordero, lNumMaxBordero, objCobrador.iCodigo, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 93569

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 93570

    Call Comando_Fechar(lComando)

    Cobrador_Le_BorderosArqRem = SUCESSO

    Exit Function

Erro_Cobrador_Le_BorderosArqRem:

    Cobrador_Le_BorderosArqRem = gErr

    Select Case gErr

        Case 93567
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93569, 93570
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BORDERO_COBRANCA ", gErr, objCobrador.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160784)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
    
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
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then Error 62427

    'Drive1.Drive = Mid(NomeDiretorio, 1, 2)

    'Dir1.Path = NomeDiretorio.Text

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case Err

        Case 62427
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", Err, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143816)

    End Select

    Exit Sub

End Sub

