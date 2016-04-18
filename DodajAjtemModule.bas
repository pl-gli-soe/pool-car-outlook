Attribute VB_Name = "DodajAjtemModule"
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

Public Sub dodaj_ajtem()

    Dim oh As AjtemHandler
    
    przygotuj_content_dla_forma oh
    pokaz_form oh
End Sub


Private Sub przygotuj_content_dla_forma(ByRef moh As AjtemHandler)


    PomieszczeniaForm.ListBoxRooms.Clear


    With moh
        Set Application.ActiveExplorer.CurrentFolder = _
           Session.GetDefaultFolder(olFolderCalendar)
           
        DoEvents
         
        Set moh.objCalendar = Session.GetDefaultFolder(olFolderCalendar)
        Set moh.objPane = Application.ActiveExplorer.NavigationPane
        Set moh.objModule = moh.objPane.Modules.GetNavigationModule(olModuleCalendar)
         
        With moh.objModule.NavigationGroups
            ' Set objGroup = .GetDefaultNavigationGroup(olMyFoldersGroup)
            
            ' not working
            ' Set objGroup = .GetDefaultNavigationGroup(olRoomsGroup)
            
            ' To use a different group
            ' msoLanguageIDPolish 1045
            ' Set objGroup = item("Pomieszczenia")
            Set moh.objGroup = .item("Rooms") ' or Pomieszczenia w zaleznosci od jezyka
                
        End With
        
        iterator = 0
        For i = 1 To moh.objGroup.NavigationFolders.Count
        
            Set moh.objNavFolder = moh.objGroup.NavigationFolders.item(i)
            
            If moh.objNavFolder.DisplayName Like "*Gliwice*SG*" Then
            
                PomieszczeniaForm.ListBoxRooms.AddItem CStr(moh.objNavFolder.DisplayName)
            End If
        Next i
        
    End With

End Sub

Private Sub pokaz_form(ByRef moh As AjtemHandler)

    PomieszczeniaForm.Show
End Sub

' ten sub pracuje pod kliknieciu w button forma pomieszczen
Public Sub stworz_nowy_ajtem_na_podstawie_wyboru(wybrana_wartosc As String)

    ' FINAL SECTION after click event
    ' =========================================================
    Dim moh As AjtemHandler
    
    
    
    With moh
        Set Application.ActiveExplorer.CurrentFolder = _
           Session.GetDefaultFolder(olFolderCalendar)
           
        DoEvents
         
        Set moh.objCalendar = Session.GetDefaultFolder(olFolderCalendar)
        Set moh.objPane = Application.ActiveExplorer.NavigationPane
        Set moh.objModule = moh.objPane.Modules.GetNavigationModule(olModuleCalendar)
         
        With moh.objModule.NavigationGroups
            ' Set objGroup = .GetDefaultNavigationGroup(olMyFoldersGroup)
            
            ' not working
            ' Set objGroup = .GetDefaultNavigationGroup(olRoomsGroup)
            
            ' To use a different group
            ' msoLanguageIDPolish 1045
            If CLng(Application.LanguageSettings.LanguageID(msoLanguageIDInstall)) = 1045 Then
                Set moh.objGroup = .item("Pomieszczenia")
            Else
            
            
                Set moh.objGroup = .item("Rooms") ' or Pomieszczenia w zaleznosci od jezyka
            End If
                
        End With
        
        iterator = 0
        For i = 1 To moh.objGroup.NavigationFolders.Count
        
            Set moh.objNavFolder = moh.objGroup.NavigationFolders.item(i)
            
            If CStr(moh.objNavFolder.DisplayName) = CStr(wybrana_wartosc) Then
            
               With moh.objNavFolder
                   .IsSelected = True
                   DoEvents
               
                   ' nm = CStr(.DisplayName)
                   
                   ' moh.items = .Folder.items
                   
                   stworz_item moh, wybrana_wartosc
                   ' moh.items.Add moh.item
                   
                   
                   Exit For
                End With
            End If
        Next i
        
    End With
    
    ' =========================================================
    

End Sub

Private Sub stworz_item(ByRef a As AjtemHandler, wybrana_wartosc As String)



    G_BODY_TXT = G_BODY_PREFIX & Chr(10) & _
        G_LBL_NM & ": {imie} {nazwisko}" & Chr(10) & _
        G_LBL_NUM & ": {#}" & Chr(10) & _
        G_LBL_DEPT & ": {wpisz tutaj dzial}" & Chr(10) & _
        G_LBL_CEL & ": {wpisz tutaj cel wyjazdu}" & Chr(10) & _
        G_LBL_TAF & ": {od} - {do}" & Chr(10) & _
        G_LBL_KM_START & ": {wpisz kilometry}" & Chr(10) & _
        G_LBL_KM_STOP & ": {wpisz kilometry}" & Chr(10) & _
        G_LBL_CMNT & ": {wpisz tutaj komentarz}" & Chr(10)



    Set a.item = a.objNavFolder.Folder.items.Add()
    
    With a.item
        .MeetingStatus = olMeeting
        .Recipients.Add CStr(wybrana_wartosc)
        .Location = CStr(wybrana_wartosc)
        .Subject = "Rezerwacja"
        ' .BodyFormat = olFormatHTML
        ' .RTFBody =
        
        .Body = CStr(G_BODY_TXT)
        
        .Display
    End With
End Sub
