Attribute VB_Name = "RoomsModule2"
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

' version 3.2
' pomieszczenia a roomy jako zmienna

' version 3.1

' dodane km start i stop

' version 3.0
' dodawanie nowego ajtemu z gotowym textem.
' nowe moduly i nowe typy wbudowane


' version 2.0
' rozszerzenie layoutu do poziomu coverage

' version 1.0
' udostepniona Asi i Dominice - tylko prosta tabela excelowa





Public Sub SelectCalendars()
    MyForm.Show
End Sub

Public Sub iSelectCalendars(dod As Date, ddo As Date)


    ' this collection items
    Dim elements(0 To SIZE_LIMIT) As Ajtem
    ' Dim kolekcja As Dictionary
    
    clear_arr_of_elements elements
    

    ' outlook handlers / variables
    Dim oh As AjtemHandler
    
    
     
    Dim i As Integer
    
    With oh
    
     
       Set Application.ActiveExplorer.CurrentFolder = _
           Session.GetDefaultFolder(olFolderCalendar)
           
       DoEvents
        
       Set oh.objCalendar = Session.GetDefaultFolder(olFolderCalendar)
       Set oh.objPane = Application.ActiveExplorer.NavigationPane
       Set oh.objModule = objPane.Modules.GetNavigationModule(olModuleCalendar)
        
       With oh.objModule.NavigationGroups
           ' Set objGroup = .GetDefaultNavigationGroup(olMyFoldersGroup)
           
           ' not working
           ' Set objGroup = .GetDefaultNavigationGroup(olRoomsGroup)
    
           ' To use a different group
           ' Set objGroup = item("Pomieszczenia")
            Set oh.objGroup = .item(G_ROOMS_STR) ' or Pomieszczenia w zaleznosci od jezyka
       End With
    
    
       iterator = 0
       For i = 1 To oh.objGroup.NavigationFolders.Count
           Set oh.objNavFolder = oh.objGroup.NavigationFolders.item(i)
           
           If oh.objNavFolder.DisplayName Like "*Gliwice*SG*" Then
           
           
               With oh.objNavFolder
                   .IsSelected = True
                   DoEvents
               
                   nm = CStr(.DisplayName)
                   
                   Set oh.items = .Folder.items
                   For Each oh.item In oh.items
                       'MsgBox "car: " & CStr(nm) & Chr(10) & _
                       '    "all day: " & item.AllDayEvent & Chr(10) & _
                       '    "start: " & CStr(item.Start) & ", end: " & CStr(item.End)
                       fill_arr_of_element iterator, elements, oh.item, nm
                       
                       
                       If iterator < SIZE_LIMIT Then
                           iterator = iterator + 1
                       End If
                   Next oh.item
               
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
       
       
       
       teraz_wsadz_wszystko_do_excela elements, dod, ddo
    
    
       Set oh.objPane = Nothing
       Set oh.objModule = Nothing
       Set oh.objGroup = Nothing
       Set oh.objNavFolder = Nothing
       Set oh.objCalendar = Nothing
       Set oh.objFolder = Nothing
    
    End With
    
    MsgBox "Gotowe!"
End Sub



Private Sub teraz_wsadz_wszystko_do_excela(ByRef a() As Ajtem, dod As Date, ddo As Date)


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
        'r.Offset(0, 1).Value = "Poczatek"
        'r.Offset(0, 2).Value = "Koniec"
        'r.Offset(0, 3).Value = "Details"
        r.Offset(0, 1).Value = "Past due"
        
        kolejna_kolumna = 2
        tmp_date = CDate(Format(dod, "yyyy-mm-dd"))
        Do
            r.Offset(0, kolejna_kolumna).Value = CStr(tmp_date)
            
            kolejna_kolumna = kolejna_kolumna + 1
            tmp_date = tmp_date + 1
        Loop Until tmp_date > ddo
        
        
        Set r = sh.Range("a2")
        For x = LBound(a) To UBound(a)
        
        
            If CStr(Trim(r)) = Trim(a(x).nm) And Trim(r) <> "" Then
                ' lecimy tutaj z juz istniejacym wpisem lub nie znalazlem nic i wpisuje w nowe
                
                ' zrob cov
                zrob_cov_dla_tego a(x), r, x
                
                Set r = sh.Range("a2")
            Else
                If Trim(r) <> "" Then
                    Set r = r.Offset(1, 0)
                End If
            End If
            
            If Trim(r) = "" Then
                r.Value = a(x).nm
                'r.Offset(0, 1).Value = "X" ' a(x).poczatek
                'r.Offset(0, 2).Value = "Y" ' a(x).koniec
                'r.Offset(0, 3).Value = "Z" ' a(x).details
                
                ' zrob cov
                zrob_cov_dla_tego a(x), r, x
                
                Set r = sh.Range("a2")

            End If
            
            
        Next x
    End With
    
    
    Set r = sh.Range("A1:ZZ" & CStr(r.Row))
    
    r.WrapText = False
    
    sh.Columns("A:ZZ").AutoFit

End Sub

Private Sub zrob_cov_dla_tego(a As Ajtem, r As Range, iterator As Variant)
    

    
    If a.nm = "" Then
        ' nop
    Else
    
    
        Dim r_poczatek As Range
        Dim r_koniec As Range
        
        Set r_poczatek = r.Offset(0, 2)
        Set r_koniec = r.Offset(0, 2)
        
        ' check if it's past due
        
    
        Do
            If CStr(Trim(r.Parent.Cells(1, r_poczatek.Column))) = "" Then
                Exit Do
            Else
                
                ' check if it's past due
                If r_poczatek.Column > 2 Then
                    If CDate(Format(a.poczatek, "yyyy-mm-dd")) < CDate(Format(r.Parent.Cells(1, r_poczatek.Column), "yyyy-mm-dd")) Then
                        Set r_poczatek = r.Offset(0, 1)
                    ElseIf CDate(Format(a.poczatek, "yyyy-mm-dd")) = CDate(Format(r.Parent.Cells(1, r_poczatek.Column), "yyyy-mm-dd")) Then
                    
                    ElseIf CDate(Format(a.poczatek, "yyyy-mm-dd")) > CDate(Format(r.Parent.Cells(1, r_poczatek.Column), "yyyy-mm-dd")) Then
                        Set r_poczatek = r_poczatek.Offset(0, 1)
                    End If
                End If
                
                
                ' check if it's past due
                If CDate(Format(a.koniec, "yyyy-mm-dd")) < CDate(Format(r.Parent.Cells(1, r_koniec.Column), "yyyy-mm-dd")) Then
                    Set r_koniec = r.Offset(0, 1)
                    Exit Do
                ElseIf CDate(Format(a.koniec, "yyyy-mm-dd")) = CDate(Format(r.Parent.Cells(1, r_koniec.Column), "yyyy-mm-dd")) Then
                    Exit Do
                ElseIf CDate(Format(a.koniec, "yyyy-mm-dd")) > CDate(Format(r.Parent.Cells(1, r_koniec.Column), "yyyy-mm-dd")) Then
                    Set r_koniec = r_koniec.Offset(0, 1)
                Else
                    Exit Do
                End If
            End If
        Loop While True
        
        Dim ir As Range
        For Each ir In r.Parent.Range(r_poczatek, r_koniec)
            If Trim(ir) = "" Then
                ir = iterator
            Else
                ir = ir & "_" & iterator
                ir.Interior.Color = RGB(240, 0, 0)
            End If
        Next ir
        
        If r_poczatek.Comment Is Nothing Then
            On Error Resume Next
            r_poczatek.AddComment CStr(a.details)
            r_poczatek.Comment.Shape.TextFrame.AutoSize = True
        End If
    
    End If
    
    
End Sub

Public Sub fill_arr_of_element(iterator As Variant, ByRef a() As Ajtem, i As AppointmentItem, nm As Variant)

    With a(iterator)
        .nm = CStr(nm)
        .poczatek = CDate(i.Start)
        .koniec = CDate(i.End)
        .details = CStr(i.Body)
        
        'Debug.Print i.Body
        'Debug.Print i.RTFBody
    End With
    
    
End Sub


Public Sub clear_arr_of_elements(ByRef a() As Ajtem)
    
    For x = LBound(a) To UBound(a)
        a(x).details = ""
        a(x).nm = ""
    Next x
End Sub
