'===============================================================================
' Module: ContractDraftGenerator
' Description: Dynamic XML/Shape injection & entity metadata retrieval
'===============================================================================

Option Explicit

Public Sub GenerateContractDraft()
    Dim wsForm As Worksheet
    Dim wsData As Worksheet
    Dim companyName As String
    Dim foundCell As Range
    
    Set wsForm = ThisWorkbook.Sheets("Formularz")
    Set wsData = ThisWorkbook.Sheets("Baza_Spolek")
    
    companyName = wsForm.Range("C4").Value
    
    If companyName = "" Then
        MsgBox "Proszę wybrać spółkę z listy!", vbExclamation, "Brak Danych"
        Exit Sub
    End If
    
    ' Search for company metadata in the database sheet
    Set foundCell = wsData.Range("A:A").Find(What:=companyName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        ' Populate metadata into the template
        wsForm.Range("C8").Value = wsData.Cells(foundCell.Row, 2).Value ' NIP
        wsForm.Range("C9").Value = wsData.Cells(foundCell.Row, 3).Value ' KRS
        wsForm.Range("C10").Value = wsData.Cells(foundCell.Row, 4).Value ' Adres
        wsForm.Range("C11").Value = wsData.Cells(foundCell.Row, 5).Value ' Nr Umowy
        
        MsgBox "Draft umowy dla spółki " & companyName & " został pomyślnie wygenerowany!", vbInformation, "Sukces"
    Else
        MsgBox "Nie znaleziono danych dla wybranej spółki.", vbCritical, "Błąd"
    End If
End Sub
