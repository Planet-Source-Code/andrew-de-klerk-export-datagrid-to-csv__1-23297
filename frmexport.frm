VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmexport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export to CSV"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmexport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdexport 
      Left            =   1230
      Top             =   3330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to CSV"
      Height          =   360
      Left            =   2820
      TabIndex        =   1
      Top             =   3390
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid dgexport 
      Bindings        =   "frmexport.frx":038A
      Height          =   3120
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   5503
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   7177
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc dcexport 
      Height          =   330
      Left            =   5070
      Top             =   45
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmexport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrConnectionString As String
Dim FExists As Boolean

Private Sub Form_Load()
    'Connect to Database ----->
    mstrConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\test.mdb;"
    dcexport.ConnectionString = mstrConnectionString
    dcexport.Visible = False
    dcexport.RecordSource = "SELECT * FROM streets"
    dcexport.Refresh


End Sub


Private Sub cmdExport_Click()
  Dim fieldnum As Integer
  Dim cellstring As String
  Dim headstring As String
  Dim daAnswer
  
 cdexport.CancelError = True
 On Error GoTo SaveErr

  With cdexport
    .DialogTitle = "Export to CSV"
    .Filter = "Excel Import File (*.csv)|*.csv"
    .FileName = "tester"
    .ShowSave
  End With
  
  FileExists (cdexport.FileName)
  If FExists = True Then
    daAnswer = MsgBox("File Exists. Overwrite?", vbYesNo + vbQuestion, "File Exists")
    If daAnswer = vbNo Then
      cmdExport_Click
    End If
  End If
     
  Open cdexport.FileName For Output As #1
  Print #1, "test Export - Programmed by InterMap (Pty) Ltd (http://www.intermap.co.za)" ' Bit of marketing
    
  dcexport.Recordset.Bookmark = dgexport.Bookmark
     
  For fieldnum = 0 To dgexport.Columns.Count - 1 'Routine for writing the header to the CSV File
    headstring = headstring & dgexport.Columns(fieldnum).Caption & ","
  Next
  Print #1, headstring

  Do While dcexport.Recordset.EOF = False  'For each row in the datacontrol
    For fieldnum = 0 To dcexport.Recordset.Fields.Count - 1
        cellstring = cellstring & dcexport.Recordset.Fields(fieldnum).Value & "," ' comma is for the csv format
    Next
    Print #1, cellstring
    cellstring = "'"

    dcexport.Recordset.MoveNext
  Loop
  Close #1

SaveErr:
    If Err <> 32755 Then ' 32755 : Cancel was selected
    End If
    Exit Sub

End Sub

'I claim no credit for this routine. I discovered this on PSC a long time ago, and
'it has been part of my applications ever since - short and neat!

Function FileExists(ByVal FileName As String)

   Dim Exists As Integer
   
   On Local Error Resume Next 'If some problem continue, code handles problems inherintly
   Exists = Len(Dir(FileName$)) 'Dir returns either a null string (len 0) or a filename
   On Local Error GoTo 0
 If Exists = 0 Then 'Null string?
    FileExists = False
    FExists = False
Else
    FileExists = True
    FExists = True
End If
End Function

