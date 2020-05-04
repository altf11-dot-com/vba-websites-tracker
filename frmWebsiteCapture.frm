VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWebsiteCapture 
   Caption         =   "Website Input"
   ClientHeight    =   4150
   ClientLeft      =   40
   ClientTop       =   150
   ClientWidth     =   14590
   OleObjectBlob   =   "frmWebsiteCapture.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWebsiteCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    'click event for save button
    'i.e. this code will run when save button is clicked
    '"CommandButton1 is the name of the object, which
    'can be changed in the object properties
    
    Me.Hide 'the Me object is the current form, hide from view
    'NOTE the data in the fields persists after the form is hidden
    
    'capture the data to variables
    'NOTE whether or not to create variables is a design decision
    Dim url As String, txt As String, un As String, pw As String, cmt As String
    url = TextBox1.Value
    txt = TextBox2.Value
    un = TextBox3.Value
    pw = TextBox4.Value
    cmt = TextBox5.Value

    Dim requiredFieldIsMissing As Long 'since 0, 1 or 2, could use byte or integer
    'these one liner IF statements test for empty values
    'in effect, turns both url and txt into required items
    'if either is empty the program skips adding the row
    If txt = vbNullString Then requiredFieldIsMissing = 2
    If url = vbNullString Then requiredFieldIsMissing = 1
    If requiredFieldIsMissing > 0 Then
        MsgBox "Required field missing:" & vbLf & IIf(requiredFieldIsMissing = 1, "URL", "Text Display"), vbOKOnly, "Returning to Main..."
        GoTo Exit_Sub_
    Else
        'two more subroutine calls finish the job
        'add a row, and if that's succesful put data in it
        If GetNewDataRow(txt) Then
            PutNewData url, txt, un, pw, cmt
        End If
    End If
   
Exit_Sub_:
    
    'clears the form and removes it from memory
    'if we didn't do this, we would have to write code
    'to clear the form before using it again
    Unload Me
    
End Sub

Private Sub CommandButton2_Click()
    'click event for the cancel button
    TextBox1.Value = vbNullString 'change event exits when box 1 is null
    Me.Hide
    Unload Me
End Sub

Private Sub TextBox1_Change()
    'copy URL to Text Display for editing convenience
    
    'store our URL so we can work on it
    Dim s As String
    s = TextBox1.Value
    
    If s = vbNullString Then Exit Sub
    
    'check and warn in case the URL contains backslash characters
    'why do we care? https://www.howtogeek.com/181774/why-windows-uses-backslashes-and-everything-else-uses-forward-slashes/
    If InStr(s, "\") > 0 Then
        If MsgBox("URL contains backslash", vbOKCancel, "WARNING") = vbCancel Then Unload Me
    End If
    
    'we'll remove this list of typical prefixes from our URL
    Dim checkStrings As String: checkStrings = "https://,http://,www.,.com,.org,.edu"
    
    'v and vs sets up a loop through an array
    Dim vs As Variant 'vs is an array
    vs = Split(checkStrings, ",")
    Dim v As Variant 'v will be each element of the array
    For Each v In vs 'delete each of our checkstrings if present
        s = Replace(s, v, "")
    Next v
    
    'check for a forward slash and if present, chop the string there
    Dim pos As Long
    pos = InStr(s, "/") 'pos is the position of an fslash if it exists in the string, zero if not
    If pos > 1 Then s = Left(s, pos - 1) 'chop the string at the first forward slash
    
    'put the result into the Text Display so it's already somewhat edited
    TextBox2.Value = WorksheetFunction.Proper(s) 'some familiar functions don't exist in VBA
    
End Sub

Private Function GetNewDataRow(txt As String) As Boolean
    'ASSUMED column labels (aka headers) in row 1
    'maybe there's an existing list, maybe not
   
    WebsitesPWs.Activate 'confirm activesheet
    
    'filtered data may result in false bottom row
    'so just in case filtering is on, turn it off
    ActiveSheet.AutoFilterMode = False
    
    'make a note of the bottom row by putting into variable bottRow
    'bottRow is used later to decide when to exit Do...Loop
    Dim bottRow As Long 'declare a procedure scoped variable
    bottRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    [B2].Activate 'activate our starting point

    'if .End(xlup).Row returns 1)
    'then the list is empty
    'we don't need to scan the list
    If bottRow = 1 Then
        GetNewDataRow = True
        Exit Function
    End If
    
    'but if the bottom row is greater than 2
    'there are items to scan, loop through the list
    'until the active cell contains a value greater
    'than our label in variable txt
    Do
        With ActiveCell 'With statement binds to its object ...
        'in this case the active cell (which is by definition a Range object)
        
            If .Row > bottRow Then Exit Do 'avoid infinite loops!
            
            If .Formula = txt Then 'if our text label already exists
            
                'issue the system sound, display alert and exit
                Beep
                If MsgBox(txt & " already exists", vbOKCancel, "WARNING") = vbCancel Then Exit Function
                
            End If
            
            'test our value against the existing list
            'ASSUMES SORTED DATA
            'Ucase function removes case sensitivity
            
            If UCase(txt) > UCase(.Formula) Then
            'if the active cell contains a value greater than our Text Display
            
                .Offset(1, 0).Activate 'move to (.Activate) the next row down
                'NOTE -- this usage of .Activate requires
                're-invoking the With statement,
                'so no more dotted references before end with
            
            Else
                            
                'otherwise, we've activated the desired row ...
                'where we want to insert a blank row
                'so exit the scanning Do...Loop
                Exit Do
            
            End If 'every if has an End If (unless its just a one liner)
            
        End With 'every With has and End With
    
    Loop
    
    'insert a row
    Rows(ActiveCell.Row).Insert 'insert a blank row and push the rest down by default
    GetNewDataRow = True

End Function

Private Sub PutNewData(url As String, txt As String, un As String, pw As String, cmt As String)

    WebsitesPWs.Activate
    
    'anchor a hyperlink object in our active cell
    ActiveSheet.Hyperlinks.Add Anchor:=ActiveCell, Address:=url, TextToDisplay:=txt
    
    'put the rest of our captured data in the rest of the row
    With ActiveCell
        'in this case we bind to a Range object, i.e. Activecell
    
        'search msdn vba Range.Offset
        .Offset(0, -1).Value = Now 'row offset zero is current row
        .Offset(0, -1).NumberFormat = "m/d/yy;@"
        .Offset(0, 1).Value = un 'column offsets -1,1,2,3
        .Offset(0, 2).Value = pw 'i.e. one col to the left (-1)
        .Offset(0, 3).Value = cmt 'and three (1,2,3) to the right
        
    End With
    Columns.EntireColumn.AutoFit
    
End Sub

Private Sub UserForm_Click()

End Sub
