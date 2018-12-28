Attribute VB_Name = "Main"
'Written in 2015 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit
  
Public Const macroName As String = "EditDrawingsProp"
Public Const macroSection As String = "Main"

Public Const draftProp As String = "Начертил"
Public Const changeProp As String = "Изменение"
Public Const standardFile As String = "Чертежный стандарт"

Dim swApp As Object

Sub Main()
    Set swApp = Application.SldWorks
    Form.dirBox.value = GetDirOfActiveDoc
    Form.draftChk.Caption = draftProp
    Form.stdChk.Caption = standardFile
    Form.draftBox.value = GetSetting2(draftProp)
    Form.stdBox.value = GetSetting2(standardFile)
    EnableChk Form.draftChk, Form.draftBox
    EnableChk Form.stdChk, Form.stdBox
    Form.Show
End Sub

Function GetStandardFilename() As String
    Dim fileOptions As Long
    Dim fileConfig As String
    Dim fileDispName As String
    
    GetStandardFilename = swApp.GetOpenFileName("Выберите чертежный стандарт", "", _
                                                "Drafting Standard (*.sldstd)|*.sldstd", _
                                                fileOptions, fileConfig, fileDispName)
End Function

Function GetDirOfActiveDoc() As String
    Dim doc As ModelDoc2
    
    Set doc = swApp.ActiveDoc
    If Not doc Is Nothing Then
        Dim fso As New FileSystemObject
        GetDirOfActiveDoc = fso.GetParentFolderName(doc.GetPathName) + "\"
    End If
End Function

Sub EnableChk(chk As CheckBox, box As TextBox)
    box.Enabled = chk.value
End Sub

Function Execute() 'mask for button
    Dim file_ As Variant
    Dim file As Object
    Dim path As String
    Dim folder As Object
    Dim err As swFileLoadError_e
    Dim wrn As swFileLoadWarning_e
    Dim filename As String
    Dim doc As ModelDoc2
    Dim err2 As swActivateDocError_e
    Dim stdFile As String
    Dim msgres As Integer
    Dim changeStd As Boolean
    
    path = Form.dirBox.value
    IsFileExist path
    If Form.stdChk.value Then
        stdFile = Form.stdBox.value
        changeStd = IsFileExist(stdFile)
    Else
        changeStd = False
    End If
    
    Set folder = CreateObject("Scripting.FileSystemObject").GetFolder(path)
    For Each file_ In folder.Files
        Set file = file_
        filename = LCase(file.path)
        If InStr(filename, "slddrw") > 0 And InStr(filename, "~$") = 0 Then
            Set doc = swApp.OpenDoc6(filename, swDocDRAWING, swOpenDocOptions_Silent, "", err, wrn)
            
            ''' job
            If Form.draftChk.value Then
                doc.Extension.CustomPropertyManager("").Add3 draftProp, swCustomInfoText, Form.draftBox.value, swCustomPropertyDeleteAndAdd
            End If
            If Form.changeChk.value Then
                doc.Extension.CustomPropertyManager("").Add3 changeProp, swCustomInfoText, "", swCustomPropertyDeleteAndAdd
            End If
            If changeStd Then
                ReloadStandard doc.Extension, stdFile
            End If
            If Form.chkStdUnits.value Then
                SetStandardUnits doc
            End If
            If Form.radSingleDim.value Then
                SetDualDimensions doc, False
            ElseIf Form.radDualDim.value Then
                SetDualDimensions doc, True
            End If
            ''' end job
            
            SaveThisDoc doc
            If wrn <> swFileLoadWarning_AlreadyOpen Then
                swApp.CloseDoc filename
            End If
            'End
        End If
    Next
End Function

Sub SetDualDimensions(doc As ModelDoc2, value As Boolean)
    doc.Extension.SetUserPreferenceToggle swDetailingDualDimensions, swDetailingDimension, value
    doc.Extension.SetUserPreferenceToggle swDetailingShowUnitsForDualDisplay, swDetailingDimension, value
End Sub

Sub SetStandardUnits(doc As ModelDoc2) 'mask for button
    '''Основные единицы длины
    doc.Extension.SetUserPreferenceInteger swUnitSystem, swDetailingNoOptionSpecified, swUnitSystem_Custom
    doc.Extension.SetUserPreferenceInteger swUnitsLinear, swDetailingNoOptionSpecified, swMM
    doc.Extension.SetUserPreferenceInteger swUnitsLinearDecimalDisplay, swDetailingNoOptionSpecified, swDECIMAL
    doc.Extension.SetUserPreferenceInteger swUnitsLinearDecimalPlaces, swDetailingNoOptionSpecified, 2
    doc.Extension.SetUserPreferenceInteger swUnitsLinearFractionDenominator, swDetailingNoOptionSpecified, 64
    'doc.Extension.SetUserPreferenceInteger swUnitsLinearFeetAndInchesFormat, swDetailingNoOptionSpecified, False
    'doc.Extension.SetUserPreferenceInteger swUnitsLinearRoundToNearestFraction, swDetailingNoOptionSpecified, True
    
    '''Двойные единицы длины
    doc.Extension.SetUserPreferenceInteger swUnitsDualLinear, swDetailingNoOptionSpecified, swINCHES
    doc.Extension.SetUserPreferenceInteger swUnitsDualLinearDecimalDisplay, swDetailingNoOptionSpecified, swDECIMAL
    doc.Extension.SetUserPreferenceInteger swUnitsDualLinearDecimalPlaces, swDetailingNoOptionSpecified, 3
    doc.Extension.SetUserPreferenceInteger swUnitsDualLinearFractionDenominator, swDetailingNoOptionSpecified, 64
    'doc.Extension.SetUserPreferenceInteger swUnitsDualLinearRoundToNearestFraction, swDetailingNoOptionSpecified, True
    'doc.Extension.SetUserPreferenceInteger swUnitsDualLinearFeetAndInchesFormat, swDetailingNoOptionSpecified, False
    
    '''Угловые единицы
    doc.Extension.SetUserPreferenceInteger swUnitsAngular, swDetailingNoOptionSpecified, swDEGREES
    doc.Extension.SetUserPreferenceInteger swUnitsAngularDecimalPlaces, swDetailingNoOptionSpecified, 2
    
    '''Единицы массы
    doc.Extension.SetUserPreferenceInteger swUnitsMassPropLength, swDetailingNoOptionSpecified, swMM
    doc.Extension.SetUserPreferenceInteger swUnitsMassPropDecimalPlaces, swDetailingNoOptionSpecified, 2
    doc.Extension.SetUserPreferenceInteger swUnitsMassPropMass, swDetailingNoOptionSpecified, swUnitsMassPropMass_Kilograms
    doc.Extension.SetUserPreferenceInteger swUnitsMassPropVolume, swDetailingNoOptionSpecified, swUnitsMassPropVolume_Meters3
End Sub

Function SaveThisDoc(ByRef doc As ModelDoc2) As Boolean
    Dim errors As swFileSaveError_e
    Dim warnings As swFileSaveWarning_e
    SaveThisDoc = AsBool(doc.Save3(swSaveAsOptions_Silent, errors, warnings))
End Function

Function AsBool(value As Boolean) As Boolean
    AsBool = CInt(value)
End Function

Sub ReloadStandard(ext As ModelDocExtension, stdFile As String)
    ext.LoadDraftingStandard stdFile
    ext.SetUserPreferenceInteger swLineFontVisibleEdgesStyle, swDetailingNoOptionSpecified, swLineCONTINUOUS
    ext.SetUserPreferenceInteger swLineFontVisibleEdgesThickness, swDetailingNoOptionSpecified, swLW_NORMAL
    ext.SetUserPreferenceInteger swLineFontSectionLineStyle, swDetailingNoOptionSpecified, swLineCHAINTHICK
End Sub

Function IsFileExist(fullname As String) As Boolean
    IsFileExist = fullname <> "" And Dir(fullname) <> ""
    If Not IsFileExist Then
        If MsgBox("Файл " & fullname & " отсутствует. Продолжить?", vbOKCancel) = vbCancel Then
            ExitApp
        End If
    End If
End Function

Function ExitApp() 'mask for button
    Unload Form
    End
End Function
