VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ImportaLotes 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   ScaleHeight     =   3795
   ScaleWidth      =   7410
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   2745
      Picture         =   "ImportaLotesOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3150
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   3675
      Picture         =   "ImportaLotesOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3150
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação do Arquivo"
      Height          =   2850
      Left            =   105
      TabIndex        =   8
      Top             =   195
      Width           =   7200
      Begin VB.Frame Frame2 
         Caption         =   "Colunas"
         Height          =   1965
         Left            =   75
         TabIndex        =   13
         Top             =   720
         Width           =   7050
         Begin VB.ComboBox colAdi 
            Height          =   315
            ItemData        =   "ImportaLotesOcx.ctx":025C
            Left            =   1215
            List            =   "ImportaLotesOcx.ctx":0275
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   645
            Width           =   2355
         End
         Begin VB.ComboBox colItem 
            Height          =   315
            ItemData        =   "ImportaLotesOcx.ctx":028E
            Left            =   4620
            List            =   "ImportaLotesOcx.ctx":02F1
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   615
            Width           =   2355
         End
         Begin VB.ComboBox colValid 
            Height          =   315
            ItemData        =   "ImportaLotesOcx.ctx":0343
            Left            =   4635
            List            =   "ImportaLotesOcx.ctx":03A6
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1110
            Width           =   2355
         End
         Begin VB.ComboBox colFabr 
            Height          =   315
            ItemData        =   "ImportaLotesOcx.ctx":03F8
            Left            =   1215
            List            =   "ImportaLotesOcx.ctx":045B
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1110
            Width           =   2355
         End
         Begin VB.ComboBox colLote 
            Height          =   315
            ItemData        =   "ImportaLotesOcx.ctx":04AD
            Left            =   4620
            List            =   "ImportaLotesOcx.ctx":0510
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   195
            Width           =   2355
         End
         Begin VB.ComboBox colQtd 
            Height          =   315
            ItemData        =   "ImportaLotesOcx.ctx":0562
            Left            =   1230
            List            =   "ImportaLotesOcx.ctx":05C5
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1545
            Width           =   2355
         End
         Begin VB.ComboBox colProduto 
            Height          =   315
            ItemData        =   "ImportaLotesOcx.ctx":0617
            Left            =   1215
            List            =   "ImportaLotesOcx.ctx":0630
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   225
            Width           =   2355
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Adição:"
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
            Height          =   315
            Left            =   135
            TabIndex        =   20
            Top             =   690
            Width           =   1035
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Item:"
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
            Height          =   315
            Left            =   3855
            TabIndex        =   19
            Top             =   660
            Width           =   705
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Validade:"
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
            Height          =   315
            Left            =   3690
            TabIndex        =   18
            Top             =   1155
            Width           =   915
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Fabricação:"
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
            Height          =   315
            Left            =   165
            TabIndex        =   17
            Top             =   1155
            Width           =   1020
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Lote:"
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
            Height          =   315
            Left            =   3855
            TabIndex        =   16
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            TabIndex        =   15
            Top             =   1590
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Produto:"
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
            Height          =   315
            Left            =   135
            TabIndex        =   14
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.TextBox NomeArquivo 
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   315
         Width           =   4995
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6750
         TabIndex        =   9
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Arquivo:"
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
         TabIndex        =   12
         Top             =   345
         Width           =   1560
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImportaLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gcolLotes As New Collection

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_PROCESSA_ARQRETCOBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Importar Lotes"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ImportaLotes"

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

    Set gcolLotes = Nothing

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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Call NomeArquivo_Validate(bSGECancelDummy)
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 192922)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(colLotes As Collection) As Long

    Set gcolLotes = colLotes

    Trata_Parametros = SUCESSO

End Function

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Public Function XLS_Importa(ByVal sDir As String, ByVal sArq As String) As Long

Dim lErro As Long
Dim objLote As ClassRastreamentoLote
Dim objExcelApp As New ClassExcelApp
Dim sValColProd As String, sValColQtde As String, sValColLote As String
Dim sValColFabr As String, sValColValid As String
Dim sValColAdi As String, sValColItem As String
Dim bComMask As Boolean, iLinha As Integer
Dim objProduto As ClassProduto, objItem As ClassItemNF

On Error GoTo Erro_XLS_Importa

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If InStr(1, UCase(NomeArquivo.Text), UCase(".csv")) <> 0 Then
        lErro = objExcelApp.Abrir_Planilha_CSV2(sDir & sArq)
    Else
        lErro = objExcelApp.Abrir_Planilha(sDir & sArq)
    End If
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For iLinha = 2 To 10000
    
        sValColProd = objExcelApp.Obtem_Valor_Celula(iLinha, colProduto.ListIndex + 1)
        sValColLote = objExcelApp.Obtem_Valor_Celula(iLinha, colLote.ListIndex + 1)
        sValColValid = ""
        sValColFabr = ""
        sValColQtde = ""
        sValColAdi = ""
        sValColItem = ""
        
        If Len(Trim(sValColProd)) = 0 Then
            Exit For
        End If
        
        If Len(Trim(sValColProd)) > 0 And Len(Trim(sValColLote)) > 0 Then

            If colValid.ListIndex <> 1 And Len(Trim(colValid.Text)) > 0 Then
                sValColValid = objExcelApp.Obtem_Valor_Celula(iLinha, colValid.ListIndex + 1)
                If Data_Critica(sValColValid) <> SUCESSO Then gError ERRO_SEM_MENSAGEM ' 300006
            End If
    
            If colFabr.ListIndex > 0 And Len(Trim(colFabr.Text)) > 0 Then
                sValColFabr = objExcelApp.Obtem_Valor_Celula(iLinha, colFabr.ListIndex + 1)
                If Data_Critica(sValColFabr) <> SUCESSO Then gError ERRO_SEM_MENSAGEM  '300007
            End If
    
            If colQtd.ListIndex > 0 And Len(Trim(colQtd.Text)) > 0 Then
                sValColQtde = objExcelApp.Obtem_Valor_Celula(iLinha, colQtd.ListIndex + 1)
                If Valor_Critica(sValColQtde) <> SUCESSO Then gError ERRO_SEM_MENSAGEM  '300008
            End If

            If colAdi.ListIndex > 0 And Len(Trim(colAdi.Text)) > 0 Then
                sValColAdi = objExcelApp.Obtem_Valor_Celula(iLinha, colAdi.ListIndex + 1)
                If Inteiro_Critica(sValColAdi) <> SUCESSO Then gError ERRO_SEM_MENSAGEM  '300008
            End If

            If colItem.ListIndex > 0 And Len(Trim(colItem.Text)) > 0 Then
                sValColItem = objExcelApp.Obtem_Valor_Celula(iLinha, colItem.ListIndex + 1)
                If Inteiro_Critica(sValColItem) <> SUCESSO Then gError ERRO_SEM_MENSAGEM  '300008
            End If

            Set objLote = New ClassRastreamentoLote
            Set objProduto = New ClassProduto
            Set objItem = New ClassItemNF

            objProduto.sNomeReduzido = sValColProd

            lErro = CF("Produto_Le_NomeReduzido", objProduto)
            If lErro <> SUCESSO And lErro <> 26927 Then gError ERRO_SEM_MENSAGEM

            If lErro <> SUCESSO Then gError 300004
                        
            'se o produto não estiver trabalhando com rastro
            If objProduto.iRastro <> PRODUTO_RASTRO_LOTE Then gError 300005
            
            objLote.sProduto = objProduto.sCodigo
            objLote.sCodigo = sValColLote
            
            lErro = CF("RastreamentoLote_Le", objLote)
            If lErro <> SUCESSO And lErro <> 75710 Then gError ERRO_SEM_MENSAGEM

            If lErro <> SUCESSO Then
                
                objLote.dtDataEntrada = Date
                objLote.dtDataFabricacao = StrParaDate(sValColFabr)
                objLote.dtDataValidade = StrParaDate(sValColValid)
            
                lErro = CF("RastreamentoLote_Grava", objLote)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End If

            objItem.dQuantidade = StrParaDbl(sValColQtde)
            Set objLote.objInfoUsu = objItem

            objLote.iFilialCli = StrParaInt(sValColAdi)
            objLote.iFilialOP = StrParaInt(sValColItem)
            
            gcolLotes.Add objLote

        End If

    Next
   
    lErro = objExcelApp.Fechar()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    XLS_Importa = SUCESSO

    Exit Function

Erro_XLS_Importa:

    XLS_Importa = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 300004
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sValColProd)

        Case 300005
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_RASTRO", gErr, objLote.sProduto)

        Case 300006
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VALIDADE_INVALIDA", gErr)

        Case 300007
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FABRICACAO_INVALIDA", gErr)

        Case 300008
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_INVALIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157922)

    End Select
    
    Call objExcelApp.Fechar

    Call Rotina_Erro(vbOKOnly, "Ocorreu um erro na importação da linha " & CStr(iLinha) & ". Favor verificar os dados.", gErr)

    Exit Function

End Function

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim sNomeArq As String
Dim sNomeDir As String
Dim iPos As Integer
Dim iPosAnt As Integer

On Error GoTo Erro_BotaoOK_Click

    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 192919
    If colProduto.ListIndex = -1 Then gError 300002
    If colLote.ListIndex = -1 Then gError 300003
    'If colAdi.ListIndex = -1 Then gError 300004
    'If colItem.ListIndex = -1 Then gError 300005
    
    'NomeArquivo.Text
    iPos = 1
    Do While iPos <> 0
        iPosAnt = iPos
        iPos = InStr(iPosAnt + 1, NomeArquivo.Text, "\")
    Loop
    
    sNomeDir = left(NomeArquivo.Text, iPosAnt)
    sNomeArq = Mid(NomeArquivo.Text, iPosAnt + 1)
    
    lErro = XLS_Importa(sNomeDir, sNomeArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr
        
        Case 192919
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)

        Case 300002
            Call Rotina_Erro(vbOKOnly, "ERRO_COLUNA_PRODUTO_NAO_PREENCHIDA", gErr)

        Case 300003
            Call Rotina_Erro(vbOKOnly, "ERRO_COLUNA_LOTE_NAO_PREENCHIDA", gErr)

        Case 300004
            'Call Rotina_Erro(vbOKOnly, "A coluna que contém a Adição não foi informada.", gErr)

        Case 300005
            'Call Rotina_Erro(vbOKOnly, "A coluna que contém o Item não foi informada.", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192921)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurar_Click()

    On Error GoTo Erro_BotaoProcurar_Click

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "xls Files (*.xls)|*.xls|xlsx Files(*.xlsx)|*.xlsx|csv Files(*.csv)|*.csv"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen

    ' Display name of selected file

    NomeArquivo.Text = CommonDialog1.FileName

    Call NomeArquivo_Validate(bSGECancelDummy)

    Exit Sub

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub NomeArquivo_Validate(Cancel As Boolean)
'Le o cabeçalho e ajusta a posição das colunas

Dim lErro As Long
Dim objExcelApp As New ClassExcelApp
Dim sColuna As String, iColuna As Integer

On Error GoTo Erro_NomeArquivo_Validate

    colProduto.Clear
    colLote.Clear
    colFabr.Clear
    colValid.Clear
    colQtd.Clear
    colAdi.Clear
    colItem.Clear

    If Len(Trim(NomeArquivo.Text)) > 0 Then

        'Abre o excel
        lErro = objExcelApp.Abrir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        If InStr(1, UCase(NomeArquivo.Text), UCase(".csv")) <> 0 Then
            lErro = objExcelApp.Abrir_Planilha_CSV(NomeArquivo.Text)
        Else
            lErro = objExcelApp.Abrir_Planilha(NomeArquivo.Text)
        End If
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        iColuna = 1
        sColuna = objExcelApp.Obtem_Valor_Celula(1, iColuna)
        Do While Len(Trim(sColuna)) <> 0
    
            colProduto.AddItem sColuna
            colLote.AddItem sColuna
            colFabr.AddItem sColuna
            colValid.AddItem sColuna
            colQtd.AddItem sColuna
            colAdi.AddItem sColuna
            colItem.AddItem sColuna
    
            If InStr(1, UCase(sColuna), "PRODUTO") <> 0 Then colProduto.ListIndex = iColuna - 1
            If InStr(1, UCase(sColuna), "LOTE") <> 0 Then colLote.ListIndex = iColuna - 1
            If InStr(1, UCase(sColuna), "FABRICACAO") <> 0 Then colFabr.ListIndex = iColuna - 1
            If InStr(1, UCase(sColuna), "VALIDADE") <> 0 Then colValid.ListIndex = iColuna - 1
            If InStr(1, UCase(sColuna), "QTD") <> 0 Then colQtd.ListIndex = iColuna - 1
            If InStr(1, UCase(sColuna), "ADICAO") <> 0 Then colAdi.ListIndex = iColuna - 1
            If InStr(1, UCase(sColuna), "ITEM") <> 0 Then colItem.ListIndex = iColuna - 1
    
            iColuna = iColuna + 1
            sColuna = objExcelApp.Obtem_Valor_Celula(1, iColuna)
        Loop

        colFabr.AddItem " "
        colValid.AddItem " "
        colQtd.AddItem " "
        colAdi.AddItem " "
        colItem.AddItem " "
    
        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_NomeArquivo_Validate:

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 300001)

    End Select

    Call objExcelApp.Fechar

    Exit Sub

End Sub

