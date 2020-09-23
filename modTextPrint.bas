Attribute VB_Name = "modTextPrint"
'Make a user-defined-type (UDT)
Public Type tPrint
   Code As String
   Name As String
   Date As Date
   Qty As Long
   Price As Long
   Total As Long
End Type
'Declare dynamic array to get data from database
Public arrPrint() As tPrint

'This function will make a string left align in text file
Function AlignLeft(NData, CFormat) As String
  If NData > 0 Then 'if not empty string
    AlignLeft = Format(NData, CFormat)
    AlignLeft = AlignLeft + Space(Len(CFormat) - Len(AlignLeft))
  Else 'empty string
    AlignLeft = Format(NData, CFormat)
    AlignLeft = "" + Space(Len(CFormat) - 1)
  End If
End Function

'This will make a string right align (usualy just for
'text in currency or number of something or numeric data)
Function AlignRight(NData, CFormat) As String
  If NData > 0 Then
    AlignRight = Format(NData, CFormat)
    AlignRight = Space(Len(CFormat) - Len(AlignRight)) + AlignRight
  Else
    AlignRight = Format(NData, CFormat)
    AlignRight = Space(Len(CFormat) - 1) + "0"
  End If
End Function

'Check whether printer has been installed in your computer
Public Function IsPrinterInstalled() As Boolean
On Error Resume Next
Dim strDummy As String
  strDummy = Printer.DeviceName
  If Err.Number Then
     IsPrinterInstalled = False
  Else
     IsPrinterInstalled = True
  End If
End Function
