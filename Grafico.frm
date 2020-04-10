VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Grafico 
   Caption         =   "Gráfico : "
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8235
      ScaleHeight     =   450
      ScaleWidth      =   1080
      TabIndex        =   0
      Top             =   60
      Width           =   1140
      Begin VB.CommandButton BotaoSair 
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
         Left            =   585
         Picture         =   "Grafico.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
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
         Left            =   75
         Picture         =   "Grafico.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   45
         Width           =   420
      End
   End
   Begin MSChart20Lib.MSChart MSGrafico1 
      Height          =   5655
      Left            =   15
      OleObjectBlob   =   "Grafico.frx":0280
      TabIndex        =   3
      Top             =   -30
      Width           =   9405
   End
   Begin VB.PictureBox PictureBug 
      Height          =   5655
      Left            =   105
      ScaleHeight     =   5595
      ScaleWidth      =   9345
      TabIndex        =   4
      Top             =   75
      Visible         =   0   'False
      Width           =   9405
   End
End
Attribute VB_Name = "Grafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MM_ANISOTROPIC = 8

Private Type Size
        cx As Long
        cy As Long
End Type

Private Type RECT
        left As Long
        top As Long
        right As Long
        bottom As Long
End Type

Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias _
       "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_PAINT = &HF

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function PlayMetafile Lib "gdi32" Alias "PlayMetaFile" (ByVal hdc As Long, ByVal hmf As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetWindowExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Private Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ScaleViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As Size) As Long

Public iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub Form_Load()

    lErro_Chama_Tela = SUCESSO

    Exit Sub

End Sub

Private Function Tela_Preenche(objGrafico As ClassGrafico, MSChart1 As Object) As Long

Dim objItemGrafico As ClassItemGrafico
Dim lErro As Long
Dim dPercent As Double
Dim dTotal As Double
Dim objItemGraficoTotal As ClassItemGrafico
Dim Row As Integer
Dim Column As Integer

On Error GoTo Erro_Tela_Preenche

    With MSChart1
        .ChartType = objGrafico.ChartType 'Tipo de gráfico
        .ColumnCount = objGrafico.colcolItensGrafico(1).Count 'Número de colunas
        .RowCount = objGrafico.colcolItensGrafico.Count 'Número de linhas
        .Plot.PlotBase.BaseHeight = 0 'Altura da base do gráfico
        .Plot.UniformAxis = False

        dTotal = 0
    
        For Row = 1 To objGrafico.colcolItensGrafico.Count
            Column = 1
            
            For Each objItemGrafico In objGrafico.colcolItensGrafico(Row)
        
                .Column = Column
                .Row = Row
                .Data = Round(objItemGrafico.dValorColuna, 2)
                .ColumnLabel = objItemGrafico.sNomeColuna
                .Plot.SeriesCollection.Item(Column).LegendText = objItemGrafico.LegendText
        
                'Configuração da Fonte que fica acima das barras do gráfico
                .Plot.SeriesCollection(Column).DataPoints.Item(-1).DataPointLabel.VtFont.Name = "Arial"
                .Plot.SeriesCollection(Column).DataPoints.Item(-1).DataPointLabel.VtFont.Size = 6
                
                If Row > 1 And Column > 1 Then
                    .Plot.SeriesCollection(Column).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeNone
                Else
                    .Plot.SeriesCollection(Column).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
                End If
                .Plot.SeriesCollection(Column).GuideLinePen.Width = 10
        
                Column = Column + 1
    
            Next
            .RowLabel = objGrafico.colcolItensGrafico(Row).Item(1).sNomeColuna
        Next
    
       Column = 1
       .Plot.View3d.Elevation = 0
       .Plot.View3d.Rotation = 0

       'Formatação do Título
       .TitleText = objGrafico.TitleText
       
       'Tamanho do Título
'       .Title.Location.RECT.Max.X = 8000 '5321
'       .Title.Location.RECT.Max.Y = 5500 '5265
'       .Title.Location.LocationType = VtChLocationTypeTop

       'Tamanho do Gráfico 6869
'       .Plot.LocationRect.Max.X = 6869
'       .Plot.LocationRect.Max.Y = 5151

       'Localização do Gráfico
       .Plot.PlotBase.BaseHeight = 0
       
       'Tamanho da Legenda
'       .Legend.Location.RECT.Max.X = 8865
'       .Legend.Location.RECT.Max.Y = 3562
'       .Legend.Location.RECT.Min.X = 6629
'       .Legend.Location.RECT.Min.Y = 2296
    
       'Formatação do texto de Rodapé, "FootNote"
       .FootNote.TextLayout.WordWrap = True
       .FootNote = objGrafico.FootNote
       
       'tamanho do footnote
'       .FootNote.Location.RECT.Max.X = 8900
'       .FootNote.Location.RECT.Max.Y = 564
'       .FootNote.Location.RECT.Min.X = 7020
'       .FootNote.Location.RECT.Min.Y = 180
       'Alinhamento do footnote
       .FootNote.Location.LocationType = VtChLocationTypeBottomRight

        'Call MSChart1.Layout
        
    End With

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 84616

        Case Else
            
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161705)

    End Select

    Exit Function

End Function

Private Sub BotaoImprimir_Click()

    Dim rv As Long, rc As RECT, X As Double, Y As Double
    Dim HeightRatio As Double, WidthRatio As Double
    Dim hmt As Long, hdcMeta As Long, szNomeArq As String
    Dim lpwinsize As Size, lpvpsize As Size, pagewidth As Long, pageheight As Long
    
    ' Make sure picturebox is same size as the chart.
    With PictureBug
       .Height = MSGrafico1.Height
       .Width = MSGrafico1.Width
    End With
'
    'Altera a orientação do Papel
    Printer.Orientation = vbPRORLandscape

    Printer.Print " "

    Call ScalePicPreviewToPrinterInches(PictureBug, HeightRatio, WidthRatio)
    
'    Printer.PaintPicture Clipboard.GetData(), 0, 0, MSGrafico1.Width, MSGrafico1.Height
'        Printer.PaintPicture PictureBug.Picture, 0, 0, MSGrafico1.Width, MSGrafico1.Height
    
    szNomeArq = String(255, 0)
    Call GetTempFileName(".", "grf", 0, szNomeArq)
    hdcMeta = CreateMetaFile(szNomeArq)
    rv = SendMessage(MSGrafico1.hWnd, WM_PAINT, hdcMeta, 0)
    hmt = CloseMetaFile(hdcMeta)
    
    Printer.ScaleMode = vbUser  ' pixels equivalent to MM_TEXT
'    pagewidth = Printer.ScaleWidth
'    pageheight = Printer.ScaleHeight
    
    Call SetMapMode(Printer.hdc, MM_ANISOTROPIC)
    If HeightRatio > WidthRatio Then
        Y = 1000 * (HeightRatio / WidthRatio)
        X = 1000
    Else
        Y = 1000
        X = 1000 * (WidthRatio / HeightRatio)
    End If
    
    X = GetDeviceCaps(Printer.hdc, LOGPIXELSX)
    Y = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
    
'    Call SetWindowExtEx(Printer.hdc, MSGrafico1.Width, MSGrafico1.Height, lpwinsize)
'    Call SetViewportExtEx(Printer.hdc, MSGrafico1.Width, MSGrafico1.Height, lpvpsize)
    Call SetWindowExtEx(Printer.hdc, X, Y, lpwinsize)
    Call SetViewportExtEx(Printer.hdc, 1440, 1440, lpvpsize)
    Call PlayMetafile(Printer.hdc, hmt)
    
    DeleteMetaFile (hmt)
    
    'Conclui a Impressão
    Printer.EndDoc

    Exit Sub

End Sub

Private Sub BotaoSair_Click()

    Unload Me

End Sub

Public Function Trata_Parametros(objGrafico As ClassGrafico) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    lErro = Tela_Preenche(objGrafico, MSGrafico1)
    If lErro <> SUCESSO Then gError 11111
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 11111, 11112
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161706)
            
    End Select
    
    Exit Function
    
End Function

Private Function ScalePicPreviewToPrinterInches(picPreview As PictureBox, HeightRatio As Double, WidthRatio As Double) As Double

    Dim Ratio As Double ' Ratio between Printer and Picture
'    Dim LRGap As Double, TBGap As Double
    Dim PgWidth As Double, PgHeight As Double
    Dim smtemp As Long
    
    ' Get the physical page size in Inches:
    PgWidth = Printer.Width / 1440
    PgHeight = Printer.Height / 1440
    ' Find the size of the non-printable area on the printer to
    ' use to offset coordinates. These formulas assume the
    ' printable area is centered on the page:
'    smtemp = Printer.ScaleMode
'    Printer.ScaleMode = vbInches
'    LRGap = (PgWidth - Printer.ScaleWidth) / 2
'    TBGap = (PgHeight - Printer.ScaleHeight) / 2
'    Printer.ScaleMode = smtemp
    ' Scale PictureBox to Printer's printable area in Inches:
    picPreview.ScaleMode = vbInches
    ' Compare the height and with ratios to determine the
    ' Ratio to use and how to size the picture box:
    HeightRatio = PgHeight / picPreview.ScaleHeight
    WidthRatio = PgWidth / picPreview.ScaleWidth
    If HeightRatio < WidthRatio Then
        Ratio = HeightRatio
    Else
        Ratio = WidthRatio
    End If
    ScalePicPreviewToPrinterInches = Ratio
End Function



