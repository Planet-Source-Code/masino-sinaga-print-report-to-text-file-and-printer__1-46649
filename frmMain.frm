VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Print Report to Text File, by Masino Sinaga (masino_sinaga@yahoo.com)"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgOpenSave 
      Left            =   3240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3840
      Top             =   720
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   955
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   955
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   0
      Width           =   955
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox rtfLap1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   2625
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4233
            MinWidth        =   4233
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "by Masino Sinaga (masino_sinaga@yahoo.com)"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "F1=Help"
            TextSave        =   "F1=Help"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Press F1 to display help"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Category"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3704
            MinWidth        =   3704
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Day and date today"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Time today"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin ComctlLib.ProgressBar prgBar1 
      Height          =   225
      Left            =   3240
      TabIndex        =   7
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox cboRec 
      Height          =   315
      ItemData        =   "frmMain.frx":00EE
      Left            =   6120
      List            =   "frmMain.frx":00F0
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblPrinter 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer:"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblJlhRec 
      BackStyle       =   0  'Transparent
      Caption         =   "Rec per page:"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuAll 
         Caption         =   "&All Products"
      End
      Begin VB.Menu mnuDateReceived 
         Caption         =   "&Date Received..."
      End
      Begin VB.Menu mnuCode 
         Caption         =   "Product &Code..."
      End
      Begin VB.Menu mnuName 
         Caption         =   "Product &Name..."
      End
      Begin VB.Menu mnuPrice 
         Caption         =   "Product &Price..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name   : frmMain.frm
'Description : Print information from database to text file
'              and to printer by using RichTextBox control.
'Reference   : "Microsoft ActiveX Data Objects 2.0 Library".
'              (from menu: Project->References...)
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date/Time   : Tuesday, July 1, 2003
'Location    : JAKARTA, INDONESIA
'-----------------------------------------------------------

'Since now, I always remember about someone said that how
'important of using Option Explicit in every module in
'Visual Basic.

'Every variable we use in this module, must be declared first.
'That is the meaning of Option Explicit that I know. CMIIW
'(Correct Me If I'm Wrong... :P )
'This can prevent using variable without declare first,
'so we can identify the variable and the type of variable we
'use in our program. By this method, we can measure whether
'our program uses many variable or not. Well, if you want to
'declare a variable, the first thing that you know is:
'"Is it necessary to use this varible?"
'If the answer is yes, declare first, and then go to the
'procedure or function, and use it immediately.
'Don't try pending to use this variable, because you
'may forget to use it, so you'll use many unused variables.
'I hope this won't be happened.

Option Explicit

'This variable is global just for this module only
Dim cnn As ADODB.Connection
Dim adoPrint As ADODB.Recordset
Dim sLastCategory As String
Dim sDay As String
Dim P As Printer

'Button Print was clicked
Private Sub cmdPrint_Click()
  Call PrintToPrinter
End Sub

'Button Exit was clicked
Private Sub cmdExit_Click()
  Unload Me
  Set frmMain = Nothing
End Sub

'Button Save was clicked
Private Sub cmdSave_Click()
  mnuSave_Click
End Sub

Private Sub Form_Load()
Dim aDay As Variant, i As Integer
  Me.Width = 8070
  Me.Height = 6225
  'Get name of day
  aDay = Array("Sunday", "Monday", "Tuesday", "Wednesday", _
               "Thursday", "Friday", "Saturday")
  sDay = aDay(Abs(Weekday(Date) - 1))
  'Display the common information at statusbar
  With StatusBar1
       .Panels(1).Text = "Click on menu Report above..."
       .Panels(4).Text = "" & sDay & ", " & Format(Date, "dd mmmm yyyy")
       .Panels(5).Text = Format(Time, "hh:mm:ss")
  End With
  'This will list all printers have been installed in
  'the computer that using this program
  If cboPrinter.ListCount = 0 Then
     For Each P In Printers
         cboPrinter.AddItem P.DeviceName
     Next
  End If
  'Display default printer in combobox
  cboPrinter.Text = Printer.DeviceName
  'Prepare for connection to database...
  Set cnn = New ADODB.Connection
  'Using cursor location in client side
  cnn.CursorLocation = adUseClient
  'This will open connection to database access in the
  'same directory with app. Database was protected by
  'password...(masinosinaga)
  cnn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & _
           "\Data.mdb;Jet OLEDB:" & _
           "Database Password=masinosinaga;"
  'This is for displaying number of record per page
  For i = 1 To 50
    cboRec.AddItem i
  Next i
  cboRec.Text = "5"
  'This is first, so there is no lastcategory has been
  'displayed before.
  sLastCategory = ""
End Sub

'Don't forget to clear the memory from this variable
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  cnn.Close
  Set cnn = Nothing
End Sub

'This will adjust the position of controls on the form
Private Sub Form_Resize()
On Error Resume Next
'Why must be Resume Next? Because if we don't put this above,
'when we minimize the form to taskbar in Windows, an error
'would be raised: --> "Error 384: A form can't be moved or
'sized while minimized or maximized". So, we put statement
'On Error Resume Next above.

  If Me.Height <= 6225 Then Me.Height = 6225
  If Me.Width <= 8070 Then Me.Width = 8070
  
  'Adjust the position of richtextbox and progressbar control
  rtfLap1.Move 50, 50, Me.ScaleWidth - 50, Me.ScaleHeight - 1100
  prgBar1.Move 100, Me.ScaleHeight - 1000, Me.Width - 500
  
  'Adjust the position of combobox printer
  lblPrinter.Move 100, Me.ScaleHeight - 700
  cboPrinter.Move lblPrinter.Width + 100, Me.ScaleHeight - 700
  
  'Adjust the position of combobox number of record per page
  lblJlhRec.Move lblPrinter.Width + cboPrinter.Width + 200, _
                 Me.ScaleHeight - 700
  cboRec.Move lblPrinter.Width + lblJlhRec.Width + cboPrinter.Width + 200, _
                 Me.ScaleHeight - 700
  
  'Adjust the position of all buttons
  cmdSave.Move Me.ScaleWidth - (cmdSave.Width + cmdPrint.Width + cmdExit.Width + 250), _
                 Me.ScaleHeight - 700
  cmdPrint.Move Me.ScaleWidth - (cmdPrint.Width + cmdExit.Width + 250), _
                 Me.ScaleHeight - 700
  cmdExit.Move Me.ScaleWidth - (cmdExit.Width + 250), _
                 Me.ScaleHeight - 700
  
  'In order that if we size the form smaller (the width smaller
  'than the first width when form_load), the information in
  'RichTextBox still look as usual (not in a mess)
  
  'Actually, this won't work because we had managed the resize
  'of the form and its all of contents.
  rtfLap1.RightMargin = rtfLap1.Width + 8000

End Sub

Private Sub mnuContents_Click()
   'Show help
   MsgBox "1. Click on menu Report, then choose which category will be displayed." & vbCrLf & _
          "" & vbCrLf & _
          "2. All, means, all record would be displayed to screen." & vbCrLf & _
          "" & vbCrLf & _
          "3. DateReceived, means, products would be displayed" & vbCrLf & _
          "   based on date you enter to InputBox." & vbCrLf & _
          "" & vbCrLf & _
          "4. Code, means, products would be displayed " & vbCrLf & _
          "   based on the code you enter to InputBox." & vbCrLf & _
          "" & vbCrLf & _
          "5. Name, means, products would be displayed" & vbCrLf & _
          "   based on the name you enter to InputBox." & vbCrLf & _
          "" & vbCrLf & _
          "6. Price, means, products would be displayed" & vbCrLf & _
          "   based on the price you enter to InputBox." & vbCrLf & _
          "" & vbCrLf & _
          "7. Click on menu File->Save or button Save" & vbCrLf & _
          "   to save report to text file with the filename." & vbCrLf & _
          "   that you can choose." & vbCrLf & _
          "" & vbCrLf & _
          "8. Click on menu File->Print or button Print" & vbCrLf & _
          "   to print report to printer.", vbInformation, "Contents"
End Sub

'This will open text file and display it to richtextbox
Private Sub mnuOpen_Click()
On Error GoTo Cancel
   With dlgOpenSave
      .Filter = "*.txt|*.txt"
      .ShowOpen
      Open .FileName For Input As #1
        rtfLap1.Text = Input(LOF(1), 1)
      Close #1
   End With
   Exit Sub
Cancel:
   Exit Sub
End Sub

'This will print the information on richtextbox to printer
Private Sub PrintToPrinter()
Dim intAsk As Integer
On Error GoTo PrintError
  If IsPrinterInstalled = False Then
     MsgBox "There is no printer has been installed" & Chr(13) & _
            "in your computer. Please install" & Chr(13) & _
            "printer first!", vbExclamation, _
            "Printer Not Install"
     Exit Sub
  Else
  End If
  If rtfLap1.Text = "" Then
     MsgBox "There is no information is being displayed to your screen this time!" & Chr(13) & _
            "Please choose the category by clicking on menu Report above" & Chr(13) & _
            "and then click on menu File->Print or button Print.", _
            vbCritical, "No Result"
     Exit Sub
  End If
  
  'Print rtfLap1.Text to printer
  Printer.FontName = "Courier New"
  Printer.FontSize = "9"
  Printer.Print rtfLap1.Text
  Printer.EndDoc '<-- This will eject the paper till the end
                 '    of paper
  
  'If you don't want printer roll up the paper till the end,
  'you can use the following code. The printer head will be stop
  'just after printing the last line in richtexbox control
  'Here is the code:
  'Open Printer.Port For Input As #1
  '   Printer.Print rtfLap1.Text
  'Close #1

  'Are you sure the report is correct? If so,
  'clear richtexbox, if not yet, leave it...
  If MsgBox("The information in your screen has been sent to the printer." & vbCrLf & _
            "Are you sure you want to clear the information on your screen?", _
            vbInformation + vbYesNo, "Print") = vbYes Then
     rtfLap1.Text = ""
  End If
  Exit Sub
PrintError:
    MsgBox "Error number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description & "" & Chr(13) & _
           "" & Chr(13) & _
           "May be printer is still off or out of paper." & Chr(13) & _
           "Please turn on your printer now or fill in " & Chr(13) & _
           "the paper to printer. Then, try again.", _
           vbCritical, "Printer Error"
    Exit Sub
End Sub

Private Sub mnuPrint_Click()
  Call PrintToPrinter
End Sub

Private Sub mnuCode_Click()
  'sLastCategory is variable for getting the last category
  'that we had ever used in displaying data to richtexbox
  'control, so if you click on combobox record per page,
  'program will automatically call this procedure based
  'on the value of category in sLastCategory.
  sLastCategory = "PRODUCT CODE"
  Call DisplayData(sLastCategory)
End Sub

Private Sub mnuDateReceived_Click()
   sLastCategory = "DATE RECEIVED"
   Call DisplayData(sLastCategory)
End Sub

Private Sub mnuName_Click()
  sLastCategory = "PRODUCT NAME"
  Call DisplayData(sLastCategory)
End Sub

Private Sub mnuAll_Click()
  sLastCategory = "ALL PRODUCTS"
  Call DisplayData(sLastCategory)
End Sub

Private Sub mnuPrice_Click()
  sLastCategory = "PRICE"
  Call DisplayData(sLastCategory)
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

'This will display the information from database to
'richtextbox control based on the category from user.
Private Sub DisplayData(strParamCategory As String)
  
  Dim i As Integer, j As Integer, idx As Integer
  Dim intLine As Integer, SumOfCurrency As Long
  Dim intPage As String, DateNow As String
  Dim strInput As String, strCategory As String
  Dim strFileName As String, strSQL As String

  On Error GoTo PesanError

  DateNow = Format(Date, "dd mmmm yyyy")
  
  'Check category, get the SQL statement based on the category
  If strParamCategory = "ALL PRODUCTS" Then
     strSQL = "SELECT * FROM Products ORDER BY Code ASC"
  ElseIf strParamCategory = "DATE RECEIVED" Then
StartDateAgain:
     strInput = InputBox("Enter the date of received: ", _
                         "DATE RECEIVED", "28/02/2002")
     If StrPtr(strInput) = 0 Then Exit Sub
     If Not IsDate(strInput) Then
        MsgBox "Invalid date or its format!", _
               vbCritical, "Date"
        GoTo StartDateAgain
     End If
     strSQL = "SELECT * FROM Products " & _
              "WHERE DateReceived=#" & Format(strInput, "mm/dd/yyyy") & "# " & _
              "ORDER BY Code ASC"
  ElseIf strParamCategory = "PRODUCT CODE" Then
     strInput = InputBox("Enter PRODUCT CODE (you may only enter any part of it):", "PRODUCT CODE", "001")
     If StrPtr(strInput) = 0 Then Exit Sub
     strSQL = "SELECT * FROM Products " & _
              "WHERE Code LIKE '%" & strInput & "%' " & _
              "ORDER BY Code ASC"
  ElseIf strParamCategory = "PRODUCT NAME" Then
     strInput = InputBox("Enter PRODUCT NAME (you may only enter any part of it):", "PRODUCT NAME", "printer")
     If StrPtr(strInput) = 0 Then Exit Sub
     strSQL = "SELECT * FROM Products " & _
              "WHERE Name LIKE '%" & strInput & "%' " & _
              "ORDER BY Code ASC"
  ElseIf strParamCategory = "PRICE" Then
     strInput = InputBox("Enter PRODUCT PRICE (without separator):", "PRICE", "3000000")
     If StrPtr(strInput) = 0 Then Exit Sub
     If Not IsNumeric(strInput) Then
        MsgBox "Invalid PRICE or its format!", _
               vbCritical, "Invalid"
        Exit Sub
     End If
     strSQL = "SELECT * FROM Products " & _
              "WHERE Price=" & Val(strInput) & " " & _
              "ORDER BY Code ASC"
  End If
  
  'Preparing recordset variable, assign it with
  'new recordset
  Set adoPrint = New ADODB.Recordset
  
  'Open ADODB recordset
  adoPrint.Open strSQL, cnn
  
  'If there is no data we found, display message
  If adoPrint.RecordCount = 0 Then
     MsgBox "Data not found!", _
            vbCritical, "Not Found"
     Exit Sub
  End If
  
  'Display category in statusbar
  StatusBar1.Panels(3).Text = strParamCategory   'strInput
  
  strCategory = strInput
  
  'Re-declare dynamic array as number of records in adoPrint
  ReDim arrPrint(adoPrint.RecordCount)
  
  'Initialize min and max of progressbar control
  prgBar1.Min = 0
  prgBar1.Max = adoPrint.RecordCount
  
  'Initialize counter and the other variable, and richtextbox
  j = 0: SumOfCurrency = 0: rtfLap1.Text = ""
  
  strFileName = "Temp.txt" 'default temporary file name
  
  'Don't show process to user, hide it.
  rtfLap1.Visible = False
  
  'Open temporary file for this output
  Open strFileName For Output As #1
  
    'This is the first time we put data to richtextbox control,
    'and we display the header of the report
    rtfLap1.Text = _
    " THE BIG VALLEY HARDWARE COMPANY, LTD & CO" & vbCrLf & _
    " " & vbCrLf & _
    " PRODUCTS LIST" & vbCrLf & _
    " BASED ON: " & strParamCategory & " " & strCategory & " " & vbCrLf & _
    " " & UCase(sDay) & ", " & UCase(DateNow) & "" & vbCrLf & _
    "                                                              Page: 01" & vbCrLf & _
    " =====================================================================" & vbCrLf & _
    " No.  Code    Name       Date Received    Qty     Price       Total   " & vbCrLf & _
    " ---------------------------------------------------------------------"
    Print #1, rtfLap1.Text
  Close #1
  
  'After we save it to file, put it to richtexbox control
  Open strFileName For Input As #1
    rtfLap1.Text = Input(LOF(1), 1)
  Close #1
  
  i = 0: idx = 1: intLine = 0
  adoPrint.MoveFirst
  For i = 1 To adoPrint.RecordCount
  
  'I don't use Do While Not ...EOF below. I have ever read
  'from a website (sorry, I forgot the URL), that using
  'Do While... loop is slower than For ... Next loop,
  'because if we use Do While Not ...EOF, each time it loops,
  'program always checks whether EOF or not, and this will
  'make your program runs slow. Using For ... Next loop can
  'increase program speed 33% than Do While Not ...EOF.
  
  'Do While Not adoPrint.EOF
     
     'And watch this. We add a blank string each in every
     'field that we get its data from database in order that
     'to prevent an error if the value in the field is Null.
     'And if the type of the field is not string (eg numeric,
     'long, integer, etc), we add zero before the value that
     'we get from database.
     arrPrint(i).Code = "" & adoPrint.Fields("Code")
     arrPrint(i).Name = "" & adoPrint.Fields("Name")
     arrPrint(i).Date = "" & adoPrint.Fields("DateReceived")
     arrPrint(i).Qty = 0 & adoPrint.Fields("Qty")
     arrPrint(i).Price = 0 & adoPrint.Fields("Price")
     'Get the total in every record by crossing the value of
     'Qty with Price.
     arrPrint(i).Total = 0 & arrPrint(i).Qty * arrPrint(i).Price
     'Check, how many lines in every page
     If intLine = CInt(cboRec.Text) Then
        intLine = 0
        'Update counter
        idx = idx + 1
        'This will make right-align for the page-sign
        If Len(Trim(idx)) = 1 Then
           intPage = "0" & Trim(Str(idx))
        Else
           intPage = Trim(Str(idx))
        End If
        
        'Every time we put data from database to this
        'richtextbox control, we use SelStart property of
        'this control which this value we get from the
        'length of the data already exists in this richtexbox
        'control.
        rtfLap1.SelStart = Len(rtfLap1.Text)
        'So, add the data to richtextbox control
        rtfLap1.Text = rtfLap1.Text & _
        " ---------------------------------------------------------------------" & vbCrLf & _
        " Move out                                                " & _
          AlignRight(SumOfCurrency, "#,###,###,###") & "" & vbCrLf & _
        " =====================================================================" & vbCrLf & _
        " " & vbCrLf & _
        " " & vbCrLf & _
        " " & vbCrLf & _
        "                                                              Page: " & intPage & "" & vbCrLf & _
        " =====================================================================" & vbCrLf & _
        " No.  Code    Name       Date Received    Qty     Price       Total   " & vbCrLf & _
        " ---------------------------------------------------------------------" & vbCrLf & _
        " Move in (from above)                                    " & _
          AlignRight(SumOfCurrency, "#,###,###,###") & "" & vbCrLf
     End If
     
     'Save again to temporary file
     Open strFileName For Output As #1
       'Get the new start of next record in richtexbox control
       rtfLap1.SelStart = Len(rtfLap1.Text)
       'This will display every record in report (detail report)
       rtfLap1.Text = rtfLap1.Text & "" & AlignRight((j + 1), "###") & ". " & _
              AlignLeft(arrPrint(i).Code, "#####") & "  " & _
              AlignLeft(arrPrint(i).Name, "##########") & "   " & _
              arrPrint(i).Date & "  " & _
              AlignRight(arrPrint(i).Qty, "###,###") & "  " & _
              AlignRight(arrPrint(i).Price, "#,###,###") & "  " & _
              AlignRight(arrPrint(i).Total, "#,###,###,###") & "" & vbCrLf
              Print #1, rtfLap1.Text
     Close #1
     'Don't forget this...
     If adoPrint.EOF = True Then
        Exit For
     End If
     'Update counter for next record
     j = j + 1
     'Update the number of line
     intLine = intLine + 1
     'Get the summary of currency
     SumOfCurrency = SumOfCurrency + arrPrint(i).Total
     'Update progressbar value
     prgBar1.Value = prgBar1.Value + 1
     'Move to next record, process again...
     adoPrint.MoveNext
  'Loop  '<-- We don't use this. It's slower !!!  :(
  Next i '<-- We use this :)
  
  'Save it to temporary file (last saving)
  Open strFileName For Output As #1
    rtfLap1.Text = rtfLap1.Text & _
    " ---------------------------------------------------------------------" & vbCrLf & _
    " Total Products = " & AlignRight(j, "###,###") & ";      Total Sum Of Price = " & _
    AlignRight(SumOfCurrency, "#,###,###,###,###") & "" & vbCrLf & _
    " =====================================================================" & vbCrLf & _
    " " & vbCrLf & _
    " MANAGER," & String$(36, " ") & "OPERATOR," & vbCrLf & _
    " " & vbCrLf & _
    " " & vbCrLf & _
    " " & vbCrLf & _
    " ______________" & String$(30, " ") & "_______________"
    Print #1, rtfLap1.Text
  Close #1
  'Don't forget to clear memory from object variable
  Set adoPrint = Nothing
  'Well, we already finish process, display the result
  rtfLap1.Visible = True
  'Make the value of progress bar back to zero
  prgBar1.Value = 0
  Exit Sub
PesanError: 'If we got error, display the number and description
  MsgBox Err.Number & " " & Err.Description
End Sub

'This will save the information in richtexbox control
'to another file or to the exist file
Private Sub mnuSave_Click()
On Error GoTo Cancel
   With dlgOpenSave
     .CancelError = True
     .DialogTitle = "Save as text file..."
     .Filter = "*.txt|*.txt"
     .ShowSave
     If Dir(.FileName) <> "" Then
       Dim intAsk As Integer
       intAsk = MsgBox(.FileName & " already exists." & vbCrLf & _
                       "Do you want to replace it?", _
                       vbExclamation + vbYesNo, "Replace")
       If intAsk = vbNo Then
          GoTo Cancel
       End If
     End If
     Open .FileName For Output As #1
        Print #1, rtfLap1.Text
     Close #1
   End With
   Exit Sub
Cancel:
   Exit Sub
End Sub

Private Sub mnuAbout_Click()
  MsgBox "This project shows you how you can print a simple report" & vbCrLf & _
         "to text file and to printer." & vbCrLf & _
         "" & vbCrLf & _
         "Any comments and votes would be truly appreciated." & vbCrLf & _
         "" & vbCrLf & _
         "(c) Masino Sinaga (masino_sinaga@yahoo.com)", _
         vbInformation, "About"
End Sub

Private Sub Timer1_Timer()
  StatusBar1.Panels(5).Text = Format(Time, "hh:mm:ss")
End Sub

Private Sub cboRec_Click()
  If sLastCategory <> "" Then Call DisplayData(sLastCategory)
End Sub

