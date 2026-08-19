'===============================================================================
' Module: WordContractGenerator
' Description: Pure Word VBA macro with UserForm logic for 12-entity drafting
'===============================================================================

Option Explicit

' Executed when clicking "Generate" inside the Word UserForm
Public Sub ExecuteContractGeneration(ByVal entityIndex As Integer)
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Prevent UI flickering during processing
    Application.ScreenUpdating = False
    
    ' 1. Inject Header Logo based on selected entity
    Call InjectHeaderLogo(doc, entityIndex)
    
    ' 2. Populate Text Variables (Bookmarks / Content Controls)
    Call PopulateEntityData(doc, entityIndex)
    
    ' 3. Format Attachments (Ensure each attachment starts on a new page)
    Call EnforceAttachmentPageBreaks(doc)
    
    Application.ScreenUpdating = True
    MsgBox "Draft umowy został pomyślnie wygenerowany!", vbInformation, "Sukces"
End Sub

Private Sub PopulateEntityData(ByRef doc As Document, ByVal entityIndex As Integer)
    With doc.Content.Find
        .Execute FindText:="{{KRS}}", ReplaceWith:="0000123456", Replace:=2 ' wdReplaceAll
        .Execute FindText:="{{NIP}}", ReplaceWith:="1234567890", Replace:=2
        .Execute FindText:="{{REGON}}", ReplaceWith:="987654321", Replace:=2
        .Execute FindText:="{{PREZES}}", ReplaceWith:="Jan Kowalski", Replace:=2
        .Execute FindText:="{{UMOWA_OPERATOR}}", ReplaceWith:="OP/2026/08/19/001", Replace:=2
        .Execute FindText:="{{EMAIL_EFK}}", ReplaceWith:="efaktura@spolka.pl", Replace:=2
    End With
End Sub

Private Sub InjectHeaderLogo(ByRef doc As Document, ByVal entityIndex As Integer)
    Dim sec As Section
    Dim logoPath As String
    
    logoPath = ThisDocument.Path & "\Logos\logo_" & entityIndex & ".png"
    
    ' Loop through all headers and insert logo
    For Each sec In doc.Sections
        If Dir(logoPath) <> "" Then
            sec.Headers(1).Range.InlineShapes.AddPicture FileName:=logoPath, LinkToFile:=False, SaveWithDocument:=True
        End If
    Next sec
End Sub

Private Sub EnforceAttachmentPageBreaks(ByRef doc As Document)
    ' Finds "ZAŁĄCZNIK" occurrences and inserts a page break before them
    With doc.Content.Find
        .Text = "ZAŁĄCZNIK"
        Do While .Execute
            .Parent.InsertBreak Type:=7 ' 7 = wdPageBreak
            .Parent.Collapse Direction:=0 ' 0 = wdCollapseEnd
        Loop
    End With
End Sub
