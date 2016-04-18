Attribute VB_Name = "TypeModule"
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

Public Type AjtemHandler
    objPane As Outlook.NavigationPane
    objModule As Outlook.CalendarModule
    objGroup As Outlook.NavigationGroup
    objNavFolder As Outlook.NavigationFolder
    objCalendar As Folder
    objFolder As Folder
    
    items As Outlook.items
    item As Outlook.AppointmentItem
    mail_item As Outlook.MailItem
    txt As String
End Type




Public Type Ajtem
    nm As String
    poczatek As Date
    koniec As Date
    details As String
End Type
