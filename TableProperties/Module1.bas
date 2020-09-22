Attribute VB_Name = "Module1"
Option Explicit

Public Const strDatabase$ = "DatabaseName"
Public Const strServer$ = "ServerAddress"
Public Const strUID$ = "YourUserName"
Public Const strPWD$ = "YourPassword"

Public cnOpen As Boolean
Public cn As ADODB.Connection

Public Function Open_cn()
  On Error GoTo ErrHandler
  Set cn = New ADODB.Connection
  cn.Provider = "MSDASQL;DRIVER={SQL Server};SERVER=" & strServer & ";trusted_connection=no;user id= sa;password=" & strPWD & " ;database=" & strDatabase & ";"
  cn.Open
  cnOpen = True
  Exit Function
ErrHandler:
  MsgBox "Connection Error. - " & Err.Number & " - " & Err.Description, vbOKOnly
  cnOpen = False
  End
End Function

Public Function Close_cn()
  If cnOpen Then
    cn.Close
    Set cn = Nothing
    cnOpen = False
  End If
End Function

Public Function CheckForNulls(Text As Variant, R As String)
  
  'Basically replaces Null with "" .. Used to avoid 'Invalid Use of Null' Error.
  
  If IsNull(Text) Then
    CheckForNulls = R
  Else
    CheckForNulls = Text
  End If
  
End Function

Public Function PrnText(X As Single, Y As Single, Align As String, Text As String, Optional textBold As Boolean, Optional textSize As Integer, Optional FitSize As Single, Optional Ret As String, Optional SkipPageCheck As Boolean, Optional FontUnderlined As Boolean, Optional SkipTrim As Boolean) As Single
    
    'Set up the return parameter
    Ret = LCase(Ret)
    If Ret <> "x" And Ret <> "y" Then Ret = "x"
    
    'If skiptrim=false then trim the text
    If Not SkipTrim Then Text = Trim$(Text)
    
    'Set Printer Bold State
    Printer.FontBold = textBold
    
    'Set font size, if passed
    If textSize > 0 Then Printer.FontSize = textSize
    
    'If text needs to be limited to a certain size, fit it here.
    If FitSize > 0 Then Text = FittedText(Text, FitSize)

    'Set X Alignment
    Select Case LCase$(Left$(Align, 1))
        Case "l", "j" 'Left or Justified
            Printer.CurrentX = X
        Case "m", "c" 'Middle or Center
            Printer.CurrentX = (X) - (Printer.TextWidth(Text) / 2)
        Case "r" 'Right
            Printer.CurrentX = (X) - Printer.TextWidth(Text)
    End Select
    
    'Set Printer Y Coordinate
    Printer.CurrentY = Y
    
    'Set Printer Underline State
    Printer.FontUnderline = FontUnderlined
    
    'Print the Text
    Printer.Print Text
    
    'Return Appropriate Function Value AS Single.
    If Ret = "x" Then
        PrnText = X + (Printer.TextWidth(Text))
    Else
        PrnText = Printer.CurrentY
    End If
End Function

Public Function FittedText(ByVal Text As String, Size As Single) As String
    If Printer.TextWidth(Text) > Size Then
        Do Until Printer.TextWidth(Text) <= (Size) - 0.1
            Text = Left(Text, Len(Text) - 1)
        Loop
        Text = Text & "..."
    End If
    FittedText = Text
End Function
