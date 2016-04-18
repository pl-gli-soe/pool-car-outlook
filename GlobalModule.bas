Attribute VB_Name = "GlobalModule"
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

Global Const SIZE_LIMIT = 16384 ' 2^14
Global Const G_BODY_PREFIX = "REZERWACJA SAMOCHODU POOLOWEGO"

Global G_BODY_TXT


Global Const G_LBL_NM = "Imie i nazwisko kierowcy"
Global Const G_LBL_NUM = "Numer personalny"
Global Const G_LBL_DEPT = "Dzial"
Global Const G_LBL_CEL = "Cel podrozy"
Global Const G_LBL_TAF = "TAF"
Global Const G_LBL_KM_START = "Km start"
Global Const G_LBL_KM_STOP = "Km stop"
Global Const G_LBL_CMNT = "Dodatkowy komentarz"


Global Const G_ROOMS_STR = "Rooms"
