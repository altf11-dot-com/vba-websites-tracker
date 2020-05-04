Attribute VB_Name = "modWebsiteCapture"
'https://AltF11.com
'jeff@jeffbrown.us
'VBA Beginner level commenting (over-commented)

'comments help us remember (or the next person figure out)
'why we made a coding decision; eg what a variable means,
'or perhaps we use a TODO comment for items we know we need
'to return to for improvement

'generally, comments should assume the reader has
'reasonable knowledge, but here we OVERCOMMENT for VBA beginners


'require explicit declaration of variables
'invoke VBE compiler real-time checking
Option Explicit

'Public declares scope of a variable,
'which is "beyond scope" for VBA beginners
'but just means its persistent and accessible
'to all subroutines while our code is running
Public WebsitesPWs As Worksheet


Public Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "W\n14"
    'aka the entry point, we could call this procedure almost anything,
    'but using the word "Main" is a common convention
    'among computer languages
    'https://en.wikipedia.org/wiki/Entry_point
    
    'if we don't have and can't create a WebsitesPWs worksheet we're done
    'so we check, and if okay we run, otherwise pop a message and shut down
    If EnvironmentIsOkay Then
        
        Do 'Do...Loop repeats until your code says stop
            
            'trigger our form; the real work happens inside the form object
            With frmWebsiteCapture 'With statement is discussed in the form
                .TextBox1.SetFocus 'puts our insertion point in the first data entry field
                .Show 'display and pass control to our form
            End With
        
        Loop While MsgBox("Click okay to continue ...", vbOKCancel) <> vbCancel 'we'll ask the user each loop
    
    Else
        
        MsgBox "Oops, unable to set worksheet", vbOKOnly, "FATAL ERROR"
        End 'shut down everything VBA (leaves only the primary Excel Application open)
    
    End If
    
End Sub

Private Function EnvironmentIsOkay() As Boolean
    'this is primarily for copy/paste users
    'email me (look up top) if you'd like a tutorial on these concepts

    On Error GoTo Error_Handler_ 'good ol'fashioned BASIC error handling
    
    'makes sure WebsitesPWs sheet exists,
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Name = "WebsitesPWs" Then 'we found it
            Set WebsitesPWs = sht 'set our public object variable
            EnvironmentIsOkay = True 'return value
            Exit Function
        End If
    Next sht
    
    'we won't get here if it exists, so being here
    'means we need to create the sheet
    Worksheets.Add
    Set WebsitesPWs = ActiveSheet
    WebsitesPWs.Name = "WebsitesPWs"
    
    'if column B (number 2) is empty, we assume we need to create a header row
    If Cells(Rows.Count, 2).End(xlUp).Row = 1 Then Range("A1:E1").Value = Split("Date,Link,Username,Password,Comments", ",")
    
    Columns.EntireColumn.AutoFit 'make it look pretty
    EnvironmentIsOkay = True 'return value
    Exit Function
    
Error_Handler_:
    EnvironmentIsOkay = False
End Function
    

