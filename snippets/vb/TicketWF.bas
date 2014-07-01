Attribute VB_Name = "TicketWF"
Sub WFSetup()

    EnsureCategoriesExist
    

End Sub

Public Function FindCategory(catStr As String) As Outlook.category

    Dim objNamespace As Outlook.NameSpace
    Dim objCategory As Outlook.category
    Dim objCatRes As Outlook.category
    Dim categoryF As Boolean
    
    'starts not found
    categoryF = False
    
    If objNamespace.Categories.Count > 0 Then
        For Each objCategory In objNamespace.Categories
            If (objCategory.Name Like category And !categoryF) Then
                objCatRes = objCategory
                categoryF = True
            End If
        Next
    End If
    
    If categoryF Then
        FindCategory = objCatRes
    Else
        FindCategory = Nil
    End If
End Function
Sub CreateCategory(catStr As String)


End Sub

Sub EnsureCategoriesExist()
    
    Dim listCat(1 To 1) As String
    
    listCat(1) = "Error"
    
    For i = 1 To 1
        If (!FindCategory(listCat(i))) Then
            CreateCategory (listCat(i))
        End If
    Next
    
End Sub


Sub T()
    
    Dim objNamespace As Outlook.NameSpace
    Dim TaskFolder As Outlook.Folder
    
    Dim currentItem As Object
    Dim currentTask As TaskItem
      
    Dim reg As New RegExp
    
    Set objNamespace = Application.GetNamespace("MAPI")
    Set TaskFolder = objNamespace.GetDefaultFolder(olFolderTasks)
    
    reg.IgnoreCase = True
    reg.Pattern = "(tarefa [0-9]{1})"
    
    
    For Each currentItem In TaskFolder.Items
        If (currentItem.Class = olTask) Then
            Set currentTask = currentItem
            
            If (reg.Test(currentTask.ConversationTopic) = True) Then
                MsgBox (currentTask.ConversationTopic)
            End If
        End If
    Next
        

End Sub



Sub ExtractInfo(Message As Outlook.MailItem)

    Dim objRegexIM As New RegExp
    Dim objRegexSD As New RegExp
    Dim objRegexNC As New RegExp
    Dim objMatch As Match
    Dim colMatches As MatchCollection

    'Ticket Numbers
    Dim IM As String
    Dim SD As String
    Dim NC As String
    Dim R As String

    
    WFSetup
    
    'Executa as regex no email
    objRegexIM.Pattern = "IM[0-9]{8}"
    objRegexSD.Pattern = "SD[0-9]{8}"
    objRegexNC.Pattern = "(NC#[0-9]{4}|NC[0-9]{4}|[0-9]{4})"
    
    'Extrai IM
    Set colMatches = objRegexIM.Execute(Message.Body)
    If (colMatches.Count = 1) Then
        IM = (colMatches.Item(0))
    End If
    
    'Extrai SD
    Set colMatches = objRegexSD.Execute(Message.Body)
    If (colMatches.Count = 1) Then
        SD = (colMatches.Item(0))
    End If
    
    'Extrai NC
    Set colMatches = objRegexNC.Execute(Message.Body)
    If (colMatches.Count = 1) Then
        NC = (colMatches.Item(0))
    End If
    
    R = "(" + IM + "/" + SD + "/" + NC + ")"
    MsgBox (R)
    
End Sub





