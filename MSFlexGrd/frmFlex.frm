VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFlex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hierarchical FlexGrid Sample"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8235
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1290
      TabIndex        =   7
      Top             =   915
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8235
      TabIndex        =   2
      Top             =   5175
      Width           =   8235
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   330
         Left            =   4815
         TabIndex        =   8
         Top             =   90
         Width           =   1065
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   330
         Left            =   3705
         TabIndex        =   4
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   330
         Left            =   2580
         TabIndex        =   3
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   2580
         TabIndex        =   6
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   3705
         TabIndex        =   5
         Top             =   90
         Width           =   1035
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxGrid 
      Height          =   4410
      Left            =   135
      TabIndex        =   1
      Top             =   645
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   7779
      _Version        =   393216
      RowHeightMin    =   2
      WordWrap        =   -1  'True
      Redraw          =   0   'False
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hierarchical FlexGrid  Usage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   -180
      TabIndex        =   0
      Top             =   -30
      Width           =   8430
   End
End
Attribute VB_Name = "frmFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrimaryRS As ADODB.Recordset, iRowData As Long, iColIndex As Long
Dim Cn As ADODB.Connection, I As Long, LastCol As Integer, iSort As Integer
Dim iCellTop As Long, iCellLeft As Long

Private Sub cmdCancel_Click()
'Return the txtEdit to default state
With txtEdit
    .Text = ""
    .Visible = False
End With

'Change buttons .Visible props
SetButtons False

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next

'Disable form from getting focus
Me.Enabled = False

'Confirm delete action
If MsgBox("Delete current record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm...") = vbNo Then
    Me.Enabled = True
    Exit Sub
End If

'Delete record
With adoPrimaryRS
    .Filter = "RecordID = " & iRowData
    
    .Delete
    
    .Filter = ""
    .MoveFirst
    
    'Remove item from FlexiGrid
    With FlxGrid
        .RemoveItem iRowData
        .Redraw = True
    End With
    
End With
    
Me.Enabled = True
    
End Sub

Private Sub cmdEdit_Click()

'Prepare the Edit text box
With txtEdit

    'Position the textbox
    .Font = FlxGrid.Font
    .FontSize = FlxGrid.CellFontSize
    .Left = iCellLeft + 150
    .Top = iCellTop
    .Height = FlxGrid.CellHeight
    .Width = FlxGrid.CellWidth
    
    'Get the column index no
    iColIndex = Me.FlxGrid.Col
    
    'Copy column text into text box
    .Text = FlxGrid.Text
    
    'Make it visible and current
    .Visible = True
    .SetFocus
    
End With

SetButtons True

End Sub
Private Sub SetButtons(bVal As Boolean)
'This controls some buttons visible & enable properties
cmdCancel.Visible = bVal
cmdDelete.Visible = Not bVal
cmdUpdate.Visible = bVal
cmdEdit.Visible = Not bVal
cmdClose.Visible = Not bVal
FlxGrid.Enabled = Not bVal
End Sub

Private Sub cmdUpdate_Click()

Dim newValue As Variant
newValue = Me.txtEdit.Text

With adoPrimaryRS
    .Filter = "RecordID = " & iRowData
    .MoveFirst
    iColIndex = iColIndex - 1
    .Fields(iColIndex).Value = newValue
    .Update
    .Filter = ""
    .Requery
    
    With FlxGrid
        Set .DataSource = Nothing
        GridCols
        .Redraw = True
    End With
    
     
End With
cmdCancel_Click
End Sub

Private Sub FlxGrid_Click()
With Me.FlxGrid
    iRowData = .RowData(FlxGrid.RowSel)
    iCellTop = .CellTop + Me.FlxGrid.Top
    iCellLeft = .CellLeft
End With

End Sub

Private Sub FlxGrid_DblClick()
DoSort
End Sub

Private Sub Form_Load()
On Error GoTo FormLoad_Err

Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

Dim strDB As String, RecCount As Long
strDB = App.Path & "\Manager.mdb"

Set Cn = New Connection
  With Cn
    .CursorLocation = adUseClient
    .ConnectionTimeout = 10
    .Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & strDB
  End With

Set adoPrimaryRS = New Recordset
With adoPrimaryRS
    .LockType = adLockOptimistic
    .CursorType = adOpenStatic
    .Source = "select LoadingDate,DepartureDate," _
     & "VehicleDriverName,VehicleConductorName," _
     & "VehicleType,VehicleRegNo," _
     & "VehicleCapacity,ItemLoaded," _
     & "QtyLoaded,WayBillNo," _
     & "VehicleDestination,RecordID from tblVehicleParticulars"
    Set .ActiveConnection = Cn
    .Open
    
    If .RecordCount > 0 Then
        GridCols
    End If
    
End With

FlxGrid.Redraw = True

FormLoad_Exit:
Exit Sub
  
FormLoad_Err:
    MsgBox "Unexpected Error: " & Err.Description

    Resume FormLoad_Exit
    
End Sub
Private Sub GridCols()
Dim strHeader As String, fldCount As Integer
Dim RowDataNum As Long, strRowHeader As String

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Prepare the Row & Column headers
strHeader = " "

strHeader = "     |Loading    |Departure  |Driver Name  |Conductor Name|" _
     & "Vehicle Type|Reg No      |Capacity   |Item Loaded|" _
     & "Qty. Loaded|Way Bill No   |Destination|;"

'strHeader = strHeader & ";"

'Use no of records to enumerate the columns as in Excel
fldCount = adoPrimaryRS.RecordCount
strRowHeader = ""
For I = 1 To fldCount
    strRowHeader = strRowHeader & "          |" & CStr(I)
Next

strHeader = strHeader & strRowHeader
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Set FlxGrid.RowData for editing/deleting purposes
'using the RecordID field entry as RowData
With FlxGrid

    .Redraw = False
    Set .DataSource = adoPrimaryRS
    adoPrimaryRS.MoveFirst
    
    For I = FlxGrid.FixedRows To FlxGrid.Rows - 1
        If adoPrimaryRS.EOF Then Exit Sub
        RowDataNum = adoPrimaryRS!RecordID
        FlxGrid.RowData(I) = RowDataNum
        adoPrimaryRS.MoveNext
    Next
    
    'Apply header
    .FormatString = strHeader
    .Redraw = True
    .Refresh
End With

End Sub
Sub DoSort()
'Sort the columns based on clicked column index
On Error Resume Next
Dim iCol As Integer, iBandColIndex As Integer

    Dim strSort As String
    With FlxGrid
        iBandColIndex = .BandColIndex
        iCol = .Col         'Get clicked column index
        strSort = .DataField(0, iBandColIndex)  'Get the associated data field
       
        If LastCol = iCol Then  'Same column clicked
        
             If iSort = 1 Then  'Last Sort = Ascending
                .Sort = 2       'Now Sort Descending
                iSort = 2
             ElseIf iSort = 2 Then  'Last Sort = Descending
                .Sort = 1           'Now Sort Ascending
                iSort = 1
            End If
            
        ElseIf LastCol <> iCol Then  'New column clicked
            .Sort = 1                'Sort Ascending
            iSort = 1
        End If
    
    End With
    
    LastCol = iCol  'Store the index of last column clicked
    
End Sub


