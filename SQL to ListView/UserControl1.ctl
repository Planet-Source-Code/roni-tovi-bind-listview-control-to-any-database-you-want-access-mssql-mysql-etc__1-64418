VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ScaleHeight     =   2385
   ScaleWidth      =   3990
   Begin MSComctlLib.ListView lw 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#####################################################
'Bind ListView Control To *Any* Database
'Feel free to use in your app, but just let me know by leaving a comment, so i 'm
'going to submit more samples for you guys; )
'SHARING is much more important than anyting....
'Any questions are welcome at root@mutluhost.com
'#####################################################
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
Dim AutoSize As Boolean
Dim strSQL As String
Dim sConnectionString As String
Dim SQL As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Private Sub UserControl_Resize()
On Error Resume Next
lw.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
If lw.ColumnHeaders.Count > 0 Then ResizeLW lw
End Sub

Public Function BindToSQL(strSQL As String, Optional iSQL As ADODB.Connection, Optional iRs As ADODB.Recordset)
    On Error GoTo err_handler
    If iSQL Is Nothing Then
        Set iSQL = SQL
        If iSQL.State > 0 Then iSQL.Close
        iSQL.Open sConnectionString
    End If
    If iRs Is Nothing Then Set iRs = Rs
    If iRs.State > 0 Then iRs.Close
    iRs.Open strSQL, iSQL, adOpenKeyset, adLockOptimistic
    
    Dim sF As Field
    With lw
    .ListItems.Clear
    .ColumnHeaders.Clear
    For Each sF In iRs.Fields
        .ColumnHeaders.Add , , sF.Name
    Next
    Dim x As Integer, item
    Do While Not iRs.EOF
        For x = 0 To iRs.Fields.Count - 1
            If x = 0 Then
                Set item = .ListItems.Add(, , iRs(x).Value)
            Else
                item.SubItems(x) = iRs(x).Value
            End If
        Next
    iRs.MoveNext
    Loop
    
    End With
    ResizeLW lw
    Exit Function
err_handler:
    MsgBox "Error binding ListView to SQL" & vbNewLine & vbNewLine & "Error code:" & Err.Number & vbNewLine & "Error desc:" & Err.Description, vbCritical
End Function
Public Property Get Enabled() As Boolean
    Enabled = lw.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lw.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
    Set Font = lw.Font
End Property
Public Property Let Font(ByVal New_Font As Font)
    Set lw.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get ConnString() As String
    ConnString = sConnectionString
End Property
Public Property Let ConnString(ByVal New_ConnString As String)
    sConnectionString = New_ConnString
    PropertyChanged "ConnString"
End Property
Public Property Get AutoSizeColumns() As Boolean
    AutoSizeColumns = AutoSize
End Property

Public Property Let AutoSizeColumns(ByVal New_AutoSize As Boolean)
    AutoSize = New_AutoSize
    PropertyChanged "AutoSizeColumns"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        lw.Font = .ReadProperty("Font", "Tahoma")
        lw.Font = .ReadProperty("Enabled", True)
        sConnectionString = .ReadProperty("ConnString", "")
        AutoSize = .ReadProperty("AutoSizeColumns", 1)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Font", lw.Font, "Tahoma")
        Call .WriteProperty("Enabled", lw.Enabled, True)
        Call .WriteProperty("ConnString", sConnectionString, "")
        Call .WriteProperty("AutoSizeColumns", AutoSize, 1)
    End With
End Sub

'This ResizeLW Function also written totally by me!
'For the best performance and best view,
'it calls an API to resize column headers first,
'but actually it wont be a really good view if the listview width is much more bigger than your columns TOTAL width, will it ?
'Well, think about you have only 2 columns, and have a maximized form. In this case, API call will resize the columns
'but there still will be LOTS OF empty areas which disturbs me too much!
'So after this API call, my function checks if there's still usable area for the columns, and if so
'it SHARES the empty areas to columns

Private Sub ResizeLW(LV As ListView)
'On Error Resume Next
'I 'm hiding the listview during resizing columnheaders
'because it causes app to slow down on old machines (especially if the display adapter's driver is missing!)
'but you can surely disable hiding by removing the next 1 line
'any questions are welcome at root@mutluhost.com
LV.Visible = False
DoEvents
If LV.ColumnHeaders.Count < 1 Then Exit Sub
Dim totalLen As Long, C As ColumnHeader, ColumnCount As Integer, x As Integer
totalLen = 0
ColumnCount = LV.ColumnHeaders.Count
 
For Each C In LV.ColumnHeaders
    SendMessage LV.hwnd, LVM_FIRST + 30, C.Index - 1, -1
    totalLen = totalLen + C.Width
Next



If totalLen < LV.Width - 200 Then
    Dim ShareVal As Long
    ShareVal = Int((LV.Width - totalLen - 400) / ColumnCount)
    For x = 1 To ColumnCount
        LV.ColumnHeaders(x).Width = LV.ColumnHeaders(x).Width + ShareVal
    Next
End If
LV.Visible = True
LV.Refresh
End Sub

