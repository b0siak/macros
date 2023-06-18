
```
Sub littleLabels()
    ' defining some variables
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim selectedRow As Range
    Dim eanCell As Range
    Dim articleCell As Range
    Dim nameCell As Range
    
    ' check if only 1 row was selected
    If TypeName(Selection) <> "Range" Or Selection.Rows.Count <> 1 Then
        MsgBox "Please select a single row"
    Exit Sub
    End If
    
    ' names of files and sheets
    Set targetWorkbook = Workbooks("Book1.xlsm")
    Set targetSheet = targetWorkbook.Sheets("Sheet1")
    
    Set selectedRow = Selection.Rows(1)
    Set eanCell = selectedRow.Cells(1)
    Set articleCell = selectedRow.Cells(2)
    Set nameCell = selectedRow.Cells(3)
    
    ' actual copy operation
    targetSheet.Range("B1").Value = EANcheck(eanCell.Value)
    targetSheet.Range("B2").Value = articleCell.Value
    targetSheet.Range("B3").Value = nameCell.Value
    
    targetWorkbook.Activate
    targetWorkbook.Sheets("Sheet2").Select
    
    ' 1st way of printing
    
    ' Application.SendKeys "^p"
    
    ' 2nd way
    
    ' ActiveWindow.SelectedSheets.PrintOut Copies:=selectedRow.Cells(4).Value, ActivePrinter:="BB"

End Sub

' EAN check and adjustment based on length
Function EANcheck(eanValue As String) As String
    If Len(eanValue) = 13 Then
        EANcheck = Left(eanValue, 12)
    ElseIf Len(eanValue) = 12 Then
        EANcheck = "0" & Left(eanValue, 11)
    Else
        EANcheck = eanValue
    End If
End Function
```


# Printer Check



```
' Written:  August 05, 2017
' Authoer:  Leith Ross
' Summary:  Returns and array of printer names and port numbers on the user's computer.
'           The API calls in this module will work with both 64 bit and 32 bit Office running Windows 7 and higher.


Private Declare PtrSafe Function RegOpenKeyEx _
    Lib "Advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As LongPtr, _
         ByVal lpctstrSubKey As String, _
         ByVal ulOptions As Long, _
         ByVal samDesired As Long, _
         ByRef phKey As LongPtr) _
    As Long


Private Declare PtrSafe Function RegEnumValue _
    Lib "Advapi32.dll" Alias "RegEnumValueA" _
        (ByVal hKey As LongPtr, _
         ByVal dwIndex As Long, _
         ByVal lptstrValueName As String, _
         ByRef lpcchValueName As Long, _
         ByVal lpReserved As Long, _
         ByRef lpType As Long, _
         ByRef lpData As Byte, _
         ByRef lpcbData As Long) _
    As Long
    
Private Declare PtrSafe Function RegCloseKey _
    Lib "Advapi32.dll" _
        (ByVal hKey As LongPtr) _
    As Long
    
Private Declare PtrSafe Function FormatMessage _
    Lib "kernel32.dll" Alias "FormatMessageA" _
        (ByVal dwFlags As Long, _
         ByVal lpSource As Long, _
         ByVal dwMessageId As Long, _
         ByVal dwLanguageId As Long, _
         ByVal lptstrBuffer As String, _
         ByVal nSize As Long, _
         ByVal vaArguments As Any) _
    As Long
    
Private Sub DisplayError(ByVal Title As String, ByVal ErrorNumber As Long)
    
    Dim errMessage  As String
    Dim lenMessage  As Integer
    Dim msg         As String
    Dim retval      As Long
    
    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
    
        lenMessage = 260
        errMessage = String(lenMessage, Chr(0))
        
        retval = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, ErrorNumber, 0&, errMessage, lenMessage, 0&)
        If retval > 0 Then
            msg = "Run-time error '" & ErrorNumber & "':" & vbLf & vbLf
            msg = msg & Left(errMessage, retval)
            MsgBox msg, vbExclamation + vbOKOnly, Title
        End If
        
End Sub


Private Function GetPrintersAndPorts() As Variant


    Dim Data()  As Byte
    Dim datType As Long
    Dim hKey    As LongPtr
    Dim index   As Long
    Dim lenData As Long
    Dim lenName As Long
    Dim prnData As Variant
    Dim prnName As String
    Dim prnPort As Variant
    Dim retval  As Long
    Dim strEnd  As Long
    Dim SubKey  As String
    Dim Text    As String
    Dim valName As String
    
    Const HKCU                  As Long = &H80000001
    Const KEY_READ              As Long = &H20019
    Const SUCCESS               As Long = 0
    Const ERROR_MORE_DATA       As Long = 234
    Const ERROR_NO_MORE_ITEMS   As Long = 259
        
        ReDim prnData(0)
        
        SubKey = "Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"


            retval = RegOpenKeyEx(HKCU, SubKey, 0&, KEY_READ, hKey)
            If retval <> 0 Then Call DisplayError("Cannot Open Registry Key", retval): Exit Function
        
            Do
Start:          ReDim Data(65535)
                lenName = 260
                valName = String(lenName, Chr(0))
                retval = RegEnumValue(hKey, index, valName, lenName, 0&, datType, Data(0), lenData)
                    Select Case retval
                        Case SUCCESS
                        Case ERROR_MORE_DATA: If Data(0) = 0 Then GoTo Start
                        Case ERROR_NO_MORE_ITEMS: Exit Do
                        Case Else: Call DisplayError("Printer Port Registry Error", retval): Exit Do
                    End Select
                index = index + 1
                
                Text = StrConv(Data, vbUnicode)
                strEnd = InStr(1, Text, Chr(0) & Chr(0)) - 1
                If strEnd > 0 Then
                    prnName = Left(valName, lenName)
                    prnPort = Split(Text, ",")(1)
                    prnData(index - 1) = prnName & " on " & prnPort
                    ReDim Preserve prnData(index)
                End If
            Loop
        
        retval = RegCloseKey(hKey)
        
        If retval <> SUCCESS Then
            Call DisplayError("Cannot Close Registry Key", retval)
        Else
            GetPrintersAndPorts = prnData
        End If
        
End Function


Public Sub ShowPrintersAndPorts()


    Dim msg      As String
    Dim Printer  As Variant
    Dim Printers As Variant
    
        Printers = GetPrintersAndPorts
    
        For Each Printer In Printers
            msg = msg & Printer & vbLf
        Next Printer
        
        MsgBox msg, vbOKOnly, "Printer Names and Ports"
        
End Sub

```



