Attribute VB_Name = "Tools"
'Written in 2015 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Sub SaveSetting2(ByRef key As String, ByRef value As String)
    SaveSetting macroName, macroSection, key, value
End Sub

Sub SaveBoolSetting(ByRef key As String, value As Boolean)
    SaveSetting2 key, BoolToStr(value)
End Sub

Function GetSetting2(ByRef key As String) As String
    GetSetting2 = GetSetting(macroName, macroSection, key, "0")
End Function

Function GetBoolSetting(ByRef key As String) As Boolean
    GetBoolSetting = StrToBool(GetSetting2(key))
End Function

Function GetIntSetting(ByRef key As String) As Integer
    GetIntSetting = StrToInt(GetSetting2(key))
End Function

Function StrToInt(ByRef value As String) As Integer
    StrToInt = IIf(IsNumeric(value), CInt(value), 0)
End Function

Function StrToBool(ByRef value As String) As Boolean
    StrToBool = IIf(IsNumeric(value), CInt(value), False)
End Function

Function BoolToStr(value As Boolean) As String
    BoolToStr = Str(CInt(value))
End Function
