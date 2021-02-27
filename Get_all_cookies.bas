Attribute VB_Name = "Get_all_cookies"
Option Explicit

' writes the cookies (name, value and domain) to columns A, B and C
Sub GetAllCookies()
    Dim cookies
    Dim cookie As cookie
    Dim i%
    Dim seleniumDriver As New ChromeDriver
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Lapa1")
    
    ' get last used row and clean the data
    i = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    
    If i > 2 Then
        ws.Range("A3", ws.Cells(i, 3)).ClearContents
    End If
    
    Call seleniumDriver.Get(CStr(ws.Cells(1, 2)))
    seleniumDriver.Manage.DeleteAllCookies
    
    Call MsgBox("Accept cookies ant press Ok")
    
    Set cookies = seleniumDriver.Manage.cookies
    
    For i = 1 To cookies.Count
        Set cookie = cookies(i)
        ws.Cells(i + 2, 1) = cookie.Name
        ws.Cells(i + 2, 2) = CStr(cookie.Value)
        ws.Cells(i + 2, 3) = cookie.domain
    Next i
End Sub

' sets the cookies from the list (if several cookies are necessary)
' and reload the page to show that the message disappears
Sub TryToSetCookies()
    Dim cookies
    Dim cookie As cookie
    Dim i%
    Dim seleniumDriver As New ChromeDriver
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Lapa1")
    
    Call seleniumDriver.Get(CStr(ws.Cells(1, 2)))
    
    i = 3
    While ws.Cells(i, 1) <> ""
        Call seleniumDriver.Manage.AddCookie(ws.Cells(i, 1), _
            ws.Cells(i, 2), ws.Cells(i, 3))
        
        i = i + 1
    Wend
    
    Call seleniumDriver.Get(CStr(ws.Cells(1, 2)))
    
    Call MsgBox("Check the result!")
End Sub
