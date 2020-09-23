VERSION 5.00
Object = "{DABBF0C5-1ABC-11D3-BF43-A8DEB2086D5E}#27.0#0"; "eXTGrid3.ocx"
Begin VB.Form frmTest 
   Caption         =   "eXTGrid 3 Beta 1 - Test"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Chiudi Programma Di Esempio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   4
      Top             =   3900
      Width           =   5775
   End
   Begin eXTGrid3.eXTGrid Grid2 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorFixed  =   8388608
      CellPicture     =   "frmTest.frx":0000
      ColAlignment0   =   9
      ColWidth0       =   960
      FixedAlignment0 =   9
      FixedCols       =   1
      FixedCols       =   1
      FocusRect       =   1
      ForeColorFixed  =   12648384
      ForeColorSel    =   -2147483634
      RowHeight0      =   240
      BeginProperty FontEditText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      ButtonBackColor =   -2147483633
      ButtonBackColor =   -2147483633
      Redraw          =   -1  'True
      Redraw          =   -1  'True
   End
   Begin eXTGrid3.eXTGrid Grid 
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2672
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorFixed  =   8421376
      CellPicture     =   "frmTest.frx":001C
      ColAlignment0   =   1
      Cols            =   3
      Cols            =   3
      ColWidth0       =   2265
      FixedAlignment0 =   1
      FocusRect       =   1
      ForeColorFixed  =   12648447
      ForeColorSel    =   -2147483634
      FormatString    =   "<Cognome e Nome                    |<Data Di Nascita                    |>Credito Lire           "
      FormatString    =   "<Cognome e Nome                    |<Data Di Nascita                    |>Credito Lire           "
      RowHeight0      =   240
      ScrollBars      =   0
      TextArray0      =   "Cognome e Nome"
      BeginProperty FontEditText {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      ButtonBackColor =   -2147483633
      ButtonBackColor =   -2147483633
      AutoComplete    =   -1  'True
      Redraw          =   -1  'True
      Redraw          =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Griglia Verticale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   2460
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "Griglia Orizzontale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   180
      Width           =   3795
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'   E' possibile collegare la griglia ad un DB
'   in formato access utilizzando il DAO.
'   Per fare questo è necessario impostare
'   la proprietà 'DatabaseName' (in fase di progettazione
'   o di esecuzione) e la proprietà 'Recordsource'
'   (solo in fase di esecuzione)

'    Grid.RecordsetType = [2 - Snapshot]
'    Grid.RecordSource = "SELECT DISTINCT * FROM TBTest"
    
    Grid.ColEditMode 0, 41, ButtonStyleBrowse, AlphaNumeric, UpperCase
    Grid.ColEditMode 1, 10, ButtonStyleDrop, Date, ProperCase, "dddd dd mmmm yyyy"
    Grid.ColEditMode 2, 12, ButtonStyleDrop, Numeric, [No Case], "0,00.00"
        
    Grid.AddLookUp 0, 1, "Sandro Folco"
    Grid.AddLookUp 0, 1, "Francesco De Gregori"
    Grid.AddLookUp 0, 1, "Fabrizio De Andrè"
    Grid.AddLookUp 0, 1, "Alighieri Dante"
    
    
    Grid2.FixedRows = 1
    Grid2.Rows = 3
    Grid2.Cols = 2
    Grid2.FormatString = ";Cognome                              |Data Di Nascita   |Credito Lire   "
    Grid2.ColWidth(1) = 3680
    Grid2.FixedRows = 0
    Grid2.ScrollBars = flexScrollBarNone
    Grid.Rows = 1
    Grid.AddItem "Francesco De Gregori" & vbTab & "04/04/1951" & vbTab & "3200000"
    Grid.AddItem "Mario Rossi" & vbTab & "02/09/1954" & vbTab & "1150000"
    Grid.AddItem "Sandro Folco" & vbTab & "02/04/1970" & vbTab & "3410500"
End Sub

