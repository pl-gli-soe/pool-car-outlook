Attribute VB_Name = "FlatTableModul"

' FORREST SOFTWARE
' POOL CAR AS OUTLOOK ROOM
' Copyright (c) 2015 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.






Sub flatTableGen()


    ' this collection items
    Dim elements(0 To SIZE_LIMIT) As Ajtem
    ' Dim kolekcja As Dictionary
    
    clear_arr_of_elements elements
    

    ' outlook handlers / variables
    Dim objPane As Outlook.NavigationPane
    Dim objModule As Outlook.CalendarModule
    Dim objGroup As Outlook.NavigationGroup
    Dim objNavFolder As Outlook.NavigationFolder
    Dim objCalendar As Folder
    Dim objFolder As Folder
    
    Dim items As items
    Dim item As AppointmentItem
    Dim txt As String
     
    Dim i As Integer
     
    Set Application.ActiveExplorer.CurrentFolder = _
        Session.GetDefaultFolder(olFolderCalendar)
        
    DoEvents
     
    Set objCalendar = Session.GetDefaultFolder(olFolderCalendar)
    Set objPane = Application.ActiveExplorer.NavigationPane
    Set objModule = objPane.Modules.GetNavigationModule(olModuleCalendar)
     
    With objModule.NavigationGroups
        ' Set objGroup = .GetDefaultNavigationGroup(olMyFoldersGroup)
        
        ' not working
        ' Set objGroup = .GetDefaultNavigationGroup(olRoomsGroup)
 
        ' To use a different group
        ' Set objGroup = item("Pomieszczenia")
        Set objGroup = .item(G_ROOMS_STR) ' or Pomieszczenia w zaleznosci od jezyka
    End With
 
 
    iterator = 0
    For i = 1 To objGroup.NavigationFolders.Count
        Set objNavFolder = objGroup.NavigationFolders.item(i)
        
        If objNavFolder.DisplayName Like "*Gliwice*SG*" Then
        
        
            With objNavFolder
                .IsSelected = True
                DoEvents
            
                nm = CStr(objNavFolder.DisplayName)
                
                Set items = .Folder.items
                For Each item In items
                    'MsgBox "car: " & CStr(nm) & Chr(10) & _
                    '    "all day: " & item.AllDayEvent & Chr(10) & _
                    '    "start: " & CStr(item.Start) & ", end: " & CStr(item.End)
                    fill_arr_of_element iterator, elements, item, nm
                    
                    
                    If iterator < SIZE_LIMIT Then
                        iterator = iterator + 1
                    End If
                Next item
            
                .IsSelected = False
            End With
            
            
            'Select Case i
            '
            ' Enter the calendar index numbers you want to open
            '    Case 1, 3, 4
            '        objNavFolder.IsSelected = True
            '
            ' Set to True to open side by side
            '        objNavFolder.IsSideBySide = False
            '    Case Else
            '        objNavFolder.IsSelected = False
            'End Select
        End If
    Next
    
    
    
    teraz_wsadz_wszystko_do_flat_table elements
 
 
    Set objPane = Nothing
    Set objModule = Nothing
    Set objGroup = Nothing
    Set objNavFolder = Nothing
    Set objCalendar = Nothing
    Set objFolder = Nothing
    
    MsgBox "Gotowe!"
End Sub



Private Sub teraz_wsadz_wszystko_do_flat_table(ByRef a() As Ajtem)


    ' excel handlers / variables
    Dim excelapp As Excel.Application
    Dim wrbks As Excel.Workbooks
    Dim wrbk As Excel.Workbook
    Dim sh As Excel.Worksheet
    Dim r As Excel.Range
    
    
    Set excelapp = New Excel.Application
    excelapp.Visible = True
    
    With excelapp
        Set wrbk = .Workbooks.Add
        Set sh = wrbk.Sheets(1)
        
        Set r = sh.Range("a1")
        
        r.Value = "Samochod"
        r.Offset(0, 1).Value = "Poczatek"
        r.Offset(0, 2).Value = "Koniec"
        r.Offset(0, 3).Value = "Imie i Nazwisko"
        r.Offset(0, 4).Value = "#"
        r.Offset(0, 5).Value = "DEPT"
        r.Offset(0, 6).Value = "CEL"
        r.Offset(0, 7).Value = "TAF"
        r.Offset(0, 8).Value = "KM START"
        r.Offset(0, 9).Value = "KM STOP"
        r.Offset(0, 10).Value = "KOMENTARZ"
        
        Set r = sh.Range("a2")
        For x = LBound(a) To UBound(a)
            
            If Trim(a(x).nm) <> "" Then
                
                
                r.Value = a(x).nm
                r.Offset(0, 1).Value = a(x).poczatek
                r.Offset(0, 2).Value = a(x).koniec
                ' r.Offset(0, 3).Value = a(x).details
                
                ' SECTION ON DETAILS
                ' =================================================
                
                lecimy_z_dalszymi_kolumnami r.Offset(0, 3), a(x).details
                
                ' =================================================
                
                Set r = r.Offset(1, 0)
            End If
        Next x
    End With
    
    
    Set r = sh.Range("A1:E" & CStr(r.Row))
    
    r.WrapText = False
    
    sh.Columns("A:ZZ").AutoFit

End Sub

Private Sub lecimy_z_dalszymi_kolumnami(ir As Range, s As String)

    'Global Const G_LBL_NM = "Imie i nazwisko kierowcy"
    'Global Const G_LBL_NUM = "Numer personalny"
    'Global Const G_LBL_DEPT = "Dzial"
    'Global Const G_LBL_DEST = "Cel podrozy"
    'Global Const G_LBL_TAF = "TAF"
    'Global Const G_LBL_CMNT = "Dodatkowy komentarz"
    
    
    
    If s Like "*" & G_BODY_PREFIX & "*" Then
        ' SECTION trafionego template'u
        ' ==========================================
        
        s = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(s, G_LBL_NM, ""), _
            G_LBL_NUM, ""), _
            G_LBL_DEPT, ""), _
            G_LBL_DEST, ""), _
            G_LBL_TAF, ""), _
            G_LBL_CMNT, ""), _
            G_LBL_KM_START, ""), _
            G_LBL_KM_STOP, "")
        
        s = Replace(Replace(s, _
            "{", ""), _
            "}", "")
        
        's = Replace(s, _
        '    Chr(10) & Chr(10), _
        '    Chr(10))
        
        s = Replace(s, _
            ": ", _
            "")
        
        tmp = Split(s, _
            Chr(10))
        
        
        i = 0
        flaga = False
        For x = LBound(tmp) To UBound(tmp)
        
            Debug.Print tmp(x)
            If Trim(CStr(tmp(x))) Like "*" & CStr(Trim(G_BODY_PREFIX)) & "*" Then
                flaga = True
                x = x + 1
            End If
            
            If flaga Then
            
                ir.Offset(0, i) = Trim(tmp(x))
                i = i + 1
            End If

        Next x
        
        ' ==========================================
    End If
    
    
End Sub


