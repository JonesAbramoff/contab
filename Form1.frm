VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   660
      Left            =   1005
      TabIndex        =   0
      Top             =   645
      Width           =   1785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type DOCINFO
          pDocName As String
          pOutputFile As String
          pDatatype As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
          ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
         pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
         pcWritten As Long) As Long


Private Sub Command1_Click()
Dim iResult As Integer, sComando As String
Dim iTam1 As Integer, iTam2 As Integer, iLen As Integer
Dim sbuffer As String
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim MyDocInfo As DOCINFO
    

                Dim sQRCode As String
                
                sQRCode = "https://www.sefaz.rs.gov.br/NFCE/NFCE-COM.aspx?chNFe=43141006354976000149650540000086781171025455&nVersao=100&tpAmb=2&dhEmi=323031342d31302d33305431353a33303a32302d30323a3030&vNF=0.10&vICMS=0.00&digVal=682f4d6b6b366134416d6f7434346d335a386947354f354b6e50453d&cIdToken=000001&cHashQRCode=771A7CE8C50D01101BDB325611F582B67FFF36D0"
                
                iLen = Len(sQRCode) + 2
                If iLen > 255 Then
                
                    iTam1 = iLen Mod 256
                    iTam2 = iLen \ 256
                    
                Else
                    iTam1 = iLen
                    iTam2 = 0
                End If
            
                sbuffer = Chr$(27) & Chr$(106) & Chr$(49) & Chr$(27) & Chr$(129) & Chr$(iTam1) & Chr$(iTam2) & Chr$(4) & "M" & sQRCode
                
                    Dim lhPrinter As Long

    If Len(sbuffer) <> 0 Then
        
        'descarregar buffer
        lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
        If lReturn = 0 Then
            'Me.List1.AddItem ("Impressora não reconhecida.")
            
        Else
        
            MyDocInfo.pDocName = "Corporator"
            MyDocInfo.pOutputFile = vbNullString
            MyDocInfo.pDatatype = "RAW"
            lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
            Call StartPagePrinter(lhPrinter)
            lReturn = WritePrinter(lhPrinter, ByVal sbuffer, Len(sbuffer), lpcWritten)
            lReturn = EndPagePrinter(lhPrinter)
            lReturn = EndDocPrinter(lhPrinter)
            lReturn = ClosePrinter(lhPrinter)
        
        End If
        
        sbuffer = ""
    
    End If

                
End Sub
