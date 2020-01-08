Attribute VB_Name = "webpw"
'https://altf11.com
'jeff@jeffbrown.us
'VBA Beginner level commenting (over-commented)

'watch the "Websites Project" series of videos
'for the line-by-line explanation

'comments help you (or the next person) remember
'why you wrote code the way you did

'generally, comments should assume reasonable knowledge
'but here we OVERCOMMENT for beginners

Option Explicit 'requires declaration of variables

'module scope declarations
Private url As String
Private txt As String
Private usr As String
Private pwd As String
Private cmt As String

Sub webpw_Main() 'coding concept, search "entry point"
Attribute webpw_Main.VB_ProcData.VB_Invoke_Func = "W\n14"
    
    webpw_GetSiteData 'calls the subroutine by this name
    
    'these one liner IF statemetnts test for empty values
    'in effect, turns both url and txt into required items
    'if either is empty the program exits
    If url = vbNullString Then Exit Sub
    If txt = vbNullString Then Exit Sub
    
    'two more subroutine calls finish the job
    webpw_ActivateInsertionRow
    webpw_InsertData
    
End Sub

Private Sub webpw_GetSiteData()
    'show our form, get data and store in variables
    'into our module-scope variables
    
    With websiteInput 'With statement binds to its object, in this case a form
    
        'displays the bound form
        .Show
        
        '(the form "closes" via the .Hide statement in the form object
        
        'store the values from the form in variables
        url = .TextBox1.Value
        txt = .TextBox2.Value
        usr = .TextBox3.Value
        pwd = .TextBox4.Value
        cmt = .TextBox5.Value
        
    End With
End Sub

Private Sub webpw_ActivateInsertionRow()
    'ASSUMED column labels (aka headers) in row 1
    'maybe there's an existing list, maybe not

    Dim bottRow As Long 'declare a procedure scoped variable
    
    'filtered data may result in false bottom row
    'so just in case filtering is on, turn it off
    ActiveSheet.AutoFilterMode = False
    
    'make a note of the bottom row by putting into variable bottRow
    'bottRow is used later to decide when to exit Do...Loop
    bottRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    [B2].Activate 'activate our starting point

    'if .End(xlup).Row returns row 2 (or 1)\
    'then the list is empty
    'we don't need to scan the list
    If bottRow <= 2 Then Exit Sub
    
    'but if the bottom row is greater than 2
    'there are items to scan, loop through the list
    'until the active cell contains a value greater
    'than our label in variable txt
    Do
        With ActiveCell 'with statement discussed in webpw_GetSiteData
        
            If .Row > bottRow Then Exit Do 'avoid infinite loops!
            
            If .Formula = txt Then 'if our text label already exists
            
                'issue the system sound, display alert and exit
                Beep
                MsgBox txt & " already exists ...", vbOKOnly, "Whoa!"
                End 'shuts down the whole thing
                
            End If
            
            'test our value against the existing list
            'ASSUMES SORTED DATA
            'Ucase function removes case sensitivity
            
            If UCase(txt) > UCase(.Formula) Then
            'if the active cell contains a value greater
                
                'move to (.Activate) the next row down
                .Offset(1, 0).Activate 'this is where we'll insert our new row
                'NOTE -- this usage of .Activate requires
                're-invoking the With statement
            
            Else
                
                'otherwise, we're here, exit the scanning Do...Loop
                Exit Do
            
            End If 'every if has an End If (unless its just a one liner)
            
        End With 'every With has and End With
        
    Loop
End Sub

Private Sub webpw_InsertData()
    'inserts a new row and writes the data
    'captured from our form in webpw_GetSiteData
    'ASSUMES the active row is where the new row is desired
    
    Rows(ActiveCell.Row).Insert 'insert a blank row and push the rest down by default
    
    'anchor a hyperlink object in our active cell
    ActiveSheet.Hyperlinks.Add Anchor:=ActiveCell, Address:=url, TextToDisplay:=txt
    
    'put the rest of our captured data in the rest of the row
    With ActiveCell 'with statement discussed in webpw_GetSiteData
        'in this case we bind to a Range object, i.e. Activecell
    
        'search msdn vba Range.Offset
        .Offset(0, -1).Value = Now 'row offset zero is current row
        .Offset(0, 1).Value = usr 'column offsets -1,1,2,3
        .Offset(0, 2).Value = pwd 'i.e. one col to the left (-1)
        .Offset(0, 3).Value = cmt 'and three (1,2,3) to the right
        
    End With 'every With has and End With
    
End Sub
