VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   " Print Table Definitions"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame LoadingFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8385
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3225
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   8
         Top             =   1380
         Width           =   585
      End
   End
   Begin VB.Frame ButtonFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   465
      Left            =   2220
      TabIndex        =   4
      Top             =   630
      Width           =   2325
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   1215
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrintAll 
         Caption         =   "Print &All"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5310
      Top             =   6150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   5655
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   9975
      _Version        =   393217
      Indentation     =   178
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.ComboBox cmbTables 
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   5925
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table to print:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   1185
   End
   Begin VB.Label lblRefresh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh Tables"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4890
      TabIndex        =   1
      Top             =   30
      Width           =   1080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const mTableName As Integer = 2
Const mFieldName As Integer = 3
Const mFieldNullable As Integer = 10
Const mFieldType As Integer = 11
Const mMaxLength As Integer = 13

Sub PopulateTables()
  Dim RS As Recordset
  Set RS = cn.OpenSchema(adSchemaColumns)
  
  'Clear Treeview and Combobox
  TV.Nodes.Clear
  cmbTables.Clear
  
  RS.MoveFirst
  Do While Not RS.EOF
    'Scroll through the database, and add all of the fields in the entire database, not just one table.
    Call AddField(CheckForNulls(RS.Fields(mTableName), "-"), CheckForNulls(RS.Fields(mFieldName), "-"), CheckForNulls(RS.Fields(mFieldNullable), "-"), CheckForNulls(RS.Fields(mFieldType), "-"), Val(CheckForNulls(RS.Fields(mMaxLength), "-")))
    DoEvents
    RS.MoveNext
  Loop
  
End Sub

Private Sub cmbTables_Click()
  
  'When a user selects a table to print, we select it in
  'the treeview, expand its children, and ensure that it is
  'visible in the list
  
  Dim X As Node
  For Each X In TV.Nodes
    If Left(X.Key, 6) = "TABLE:" Then
      If X.Key = "TABLE:" & cmbTables.Text Then
        X.EnsureVisible
        X.Expanded = True
        X.Selected = True
      Else
        X.Expanded = False
      End If
    End If
    DoEvents
  Next
  
End Sub

Private Sub cmdPrint_Click()
  Dim CurrY As Single
  Dim N As Node
  Dim K As String
  Dim Curr As String
  Dim X As Integer
  Dim L As Integer
  Dim Page As Integer
  Dim I As Integer
  Dim Attribs(0 To 4) As String
  Dim Spot As Integer
  
  Page = 1
  
  'Set Printer Settings
  Printer.ScaleMode = vbCentimeters
  Printer.FontName = "Times New Roman"
  Printer.FontSize = 8
  
  'Heading
  CurrY = PrnText(Printer.ScaleWidth / 2, 0, "C", "FIELD DEFINITIONS FOR '" & cmbTables.Text & "'   Page 1", True, 12, , "Y", , True) + 0.6
  
  'Column Headers
  Call PrnText(0, CurrY, "L", "Field", True, 9)
  Call PrnText(4, CurrY, "L", "Type", True, 9)
  Call PrnText(7, CurrY, "L", "MaxLength", True, 9, , "Y")
  CurrY = PrnText(10, CurrY, "L", "Nullable", True, 9, , "Y") + 0.3
  
  
  Printer.Line (0, CurrY)-(Printer.ScaleWidth, CurrY)
  CurrY = CurrY + 0.15
  
  
  L = Len("Field:" & cmbTables.Text & ".")
  
  For Each N In TV.Nodes
    If Left(N.Key, L) = "Field:" & cmbTables.Text & "." Then
      
      'Parse data from Node's tag
      K = N.Tag
      Spot = InStr(1, K, "|")
      Attribs(0) = Left(K, Spot - 1)
      K = Mid(K, Spot + 1)
      Spot = InStr(1, K, "|")
      Attribs(1) = Left(K, Spot - 1)
      K = Mid(K, Spot + 1)
      Spot = InStr(1, K, "|")
      Attribs(2) = Left(K, Spot - 1)
      K = Mid(K, Spot + 1)
      Spot = InStr(1, K, "|")
      Attribs(3) = Left(K, Spot - 1)
      K = Mid(K, Spot + 1)
      Attribs(4) = Left(K, Spot - 1)
      
      
      'Print the fields
      Call PrnText(0, CurrY, "L", Attribs(1), True, 9)
      Call PrnText(4, CurrY, "L", GetTypeName(Val(Attribs(3))), , 9)
      Call PrnText(7, CurrY, "L", Attribs(4), , 9, , "Y")
      CurrY = PrnText(10, CurrY, "L", Attribs(2), , 9, , "Y")
      
      'Counter for line drawing
      I = I + 1
      
      If CurrY >= Printer.ScaleHeight - 2 Then
        'If we are at the end of the page, its time for a new one:
        
        'Increment Page Count
        Page = Page + 1
        
        'Reset line counter
        I = 0
        
        'Send 'NewPage' Command to printer
        Printer.NewPage
        
        'Reset Y-Coord
        CurrY = 0
        
        'Print Heading / Column Headers / Line again:
        CurrY = PrnText(Printer.ScaleWidth / 2, 0, "C", "FIELD DEFINITIONS FOR '" & cmbTables.Text & "'   Page " & Page, True, 12, , "Y", , True) + 0.6
        Call PrnText(0, CurrY, "L", "Field", True, 9)
        Call PrnText(4, CurrY, "L", "Type", True, 9)
        Call PrnText(7, CurrY, "L", "MaxLength", True, 9, , "Y")
        CurrY = PrnText(10, CurrY, "L", "Nullable", True, 9, , "Y") + 0.3
        Printer.Line (0, CurrY)-(Printer.ScaleWidth, CurrY)
        CurrY = CurrY + 0.15
      Else
        
        'If we're not at the end of the page, check to see if this is the 5th
        'item since the last line. if so, draw a line.
        
        If I Mod 5 = 0 Then
          CurrY = CurrY + 0.15
          Printer.Line (0, CurrY)-(Printer.ScaleWidth, CurrY)
          CurrY = CurrY + 0.15
        End If
        
      End If
      
    End If
    DoEvents
  Next
  
  Printer.EndDoc
End Sub

Function GetTypeName(vType As Integer) As String
  
  'Convert the Data Type Integer value to a string
  
  Select Case vType
    Case 130
      GetTypeName = "String"
    
    Case 6
      GetTypeName = "Currency"
    
    Case 129
      GetTypeName = "Memo"
    
    Case 4
      GetTypeName = "Single"
    
    Case 135
      GetTypeName = "Date/Time"
    
    Case 11
      GetTypeName = "Boolean"
    
    Case 2
      GetTypeName = "Sm. Integer"
    
    Case 3
      GetTypeName = "Integer"
    
    Case 131
      GetTypeName = "Numeric"
    
    Case Else
      'If there is no associated type, then display the numeric value
      GetTypeName = Trim$(Str(vType))
  
  End Select
  
End Function


Private Sub cmdPrintAll_Click()
  
  'Basically, this will select each item in the combobox
  'And then click 'Print' on each. Nothing Fancy.
  
  For X = 0 To cmbTables.ListCount - 1
    cmbTables.ListIndex = X   '<---------- Select a table
    DoEvents
    cmdPrint_Click            '<---------- Print the selected table
    DoEvents
  Next
  
End Sub

Private Sub Form_Load()
  
  'Set all controls to invisible except loading frame
  SetLoading (True)
  
  'Display Application Status
  lblStatus = "Opening SQL Connection..."
  
  'Make sure the for is visible before proceding
  Visible = True
  Do Until Visible
    DoEvents
  Loop
  
  'Set the button frame's back color to the same as the forms. The reason it is different
  'In the first place, is to make it visible at design time (easier to work with)
  ButtonFrame.BackColor = BackColor
  
  'Open Database
  Open_cn
  
  'Display Application Status
  lblStatus = "Retreiving Meta Data..."
  
  'Retrieve Table names and fields
  Call PopulateTables
  
  'Set all controls back to visible except loading frame
  SetLoading (False)
  
End Sub

Sub SetLoading(isLoading As Boolean)
  
  'Set all controls to (in)visible
  Dim con As Control
  For Each con In Me
    If TypeOf X Is ComboBox Or TypeOf X Is TreeView Or TypeOf X Is Label Or TypeOf X Is Frame Then
      con.Visible = Not (isLoading)
    End If
  Next
  
  'The loading frame always does the opposite action
  LoadingFrame.Visible = isLoading
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  'Close Database
  Close_cn
  
End Sub

Private Sub Form_Resize()
  
  'If the user resizes too small, or any other error should arise,
  'just skip the rest of the sub
  On Error GoTo Err
  
  'Resize Controls
  cmbTables.Move 60, 270, ScaleWidth - 120
  lblRefresh.Move ScaleWidth - 60 - lblRefresh.Width
  ButtonFrame.Move (ScaleWidth / 2) - (ButtonFrame.Width / 2)
  TV.Move 60, 1140, ScaleWidth - 120, ScaleHeight - 1200
  
Err:

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  'Close Database
  Close_cn
  
End Sub

Sub AddField(tName As String, fName As String, fNullable As Boolean, fType As Integer, fMaxLength As Long)
  
  'In my database, all of my tables are using propercase. So, if the table name is lower-case, then
  'it is one of SQL's tables, and Im not interested in it, so I exit the sub
  If LCase(tName) = tName Then Exit Sub
  
  'Add's the Table Node - check the sub to see why it is in its own procedure.
  Call AddTableNode(tName)
  
  Dim N As Node
  
  'Add Node with key in this format:  FIELD:Accounts.AccountNumber
  Set N = TV.Nodes.Add("TABLE:" & tName, tvwChild, "Field:" & tName & "." & fName, fName, 2)
  
  'Create a delim. string (using PIPE [|]) of field's properties to its tag
  N.Tag = tName & "|" & fName & "|" & fNullable & "|" & fType & "|" & fMaxLength
  
End Sub

Sub AddTableNode(vName As String)
  
  'Add table node. The reason this is in it's own procedure, and not placed in the 'AddField'
  'Sub, is this: I do not need to check to see if the node already exists. If the node is a dupe,
  'then a 'Key is not unique' error will be raised, and the sub will terminate on its own.
  
  Dim N As Node
  On Error GoTo Err '<------------------------------
  Set N = TV.Nodes.Add(, , "TABLE:" & vName, vName, 1)
  cmbTables.AddItem vName
  
Err: '<----------------------------------------------

End Sub

Private Sub lblRefresh_Click()
  
  'Refresh Table List
  Call PopulateTables
  
End Sub
