VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY650"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY650.cls
'//  Module         : PH
'//  Desc           : ³ëµ¿Á¶ÇÕ°£ºÎÇöÈ²
'//  ÀÛ¼ºÀÚ         : NGY
'//  DATE           : 2012.12.26
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// ±×¸®µå »ç¿ë½Ã
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PY650     As SAPbouiCOM.DataTable
'
''// ¸ÅÆ®¸¯½º »ç¿ë½Ã
'Public oMat1 As SAPbouiCOM.Matrix
'Private oDS_PH_PY650A As SAPbouiCOM.DBDataSource
'Private oDS_PH_PY650B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY650.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY650_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY650"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY650_CreateItems
    Call PH_PY650_EnableMenus
    Call PH_PY650_SetDocument(oFromDocEntry01)
'    Call PH_PY650_FormResize
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing
    Exit Sub
    
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY650_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY650_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// ¾ÆÀÌÅÛ ¼³Á¤
    '//----------------------------------------------------------------------------------------------
    
    '//»ç¾÷Àå
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// Á¢¼ÓÀÚ¿¡ µû¸¥ ±ÇÇÑº° »ç¾÷Àå ÄÞº¸¹Ú½º¼¼ÆÃ
    Call CLTCOD_Select(oForm, "CLTCOD")
    
    '//±âÁØÀÏÀÚ
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
    oForm.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    oForm.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
    
    '//Ä¿¼­¸¦ Ã¹¹øÂ° ITEMÀ¸·Î ÁöÁ¤
    oForm.ActiveItem = "CLTCOD"
    
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY650_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY650_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY650_EnableMenus()
    
    On Error GoTo PH_PY650_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//Á¦°Å
    Call oForm.EnableMenu("1284", False)    '//Ãë¼Ò
    Call oForm.EnableMenu("1293", True)     '//Çà»èÁ¦
    
    Exit Sub
    
PH_PY650_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY650_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY650_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY650_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY650_FormItemEnabled
        Call PH_PY650_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY650_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY650_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY650_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY650_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY650_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        Call oForm.EnableMenu("1281", True)     '//¹®¼­Ã£±â
        Call oForm.EnableMenu("1282", False)    '//¹®¼­Ãß°¡
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then

        Call oForm.EnableMenu("1281", False)    '//¹®¼­Ã£±â
        Call oForm.EnableMenu("1282", True)     '//¹®¼­Ãß°¡
    ElseIf (oForm.Mode = fm_OK_MODE) Then

        
        Call oForm.EnableMenu("1281", True)     '//¹®¼­Ã£±â
        Call oForm.EnableMenu("1282", True)     '//¹®¼­Ãß°¡

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY650_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY650_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "Btn01" Then
                    Call PH_PY650_Print_Report01
                End If
            ElseIf pval.BeforeAction = False Then


            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat01", "Grid01"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUID
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUID
                oLastColUID = ""
                oLastColRow = 0
            End Select
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                  
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
'                    If pval.Row > 0 Then
'                        Call oMat1.SelectRow(pval.Row, True, False)
'                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Mat01", "Grid01"
                    If pval.Row > 0 Then
                        oLastItemUID = pval.ItemUID
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case Else
                    oLastItemUID = pval.ItemUID
                    oLastColUID = ""
                    oLastColRow = 0
                End Select
            ElseIf pval.BeforeAction = False Then

            End If
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7

'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
       
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID
'                        Case "Code"
'                            '//»ç¿ø¸í ªO¾Æ¼­ È­¸é Ç¥½Ã ÇÏ±â
'                            sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" & Trim(oForm.Items("Code").Specific.Value) & "'"
'                            oRecordSet.DoQuery sQry
'                            oForm.Items("CodeName").Specific.String = Trim(oRecordSet.Fields("U_FullName").Value)
                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
               'oMat1.LoadFromDataSource

                Call PH_PY650_FormItemEnabled
                Call PH_PY650_AddMatrixRow
                
            End If
'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
'                Set oDS_PH_PY650A = Nothing
'                Set oDS_PH_PY650B = Nothing
                
                'Set oMat1 = Nothing
                'Set oGrid1 = Nothing
                
            End If
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
            If pval.BeforeAction = True Then
        
            ElseIf pval.BeforeAction = False Then
                
            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
            If pval.BeforeAction = True Then
            
            ElseIf pval.Before_Action = False Then

            End If
'----------------------------------------------------------
        Case et_PICKER_CLICKED '//37
            
'----------------------------------------------------------
        Case et_GRID_SORT '//38
            
'----------------------------------------------------------
        Case et_Drag '//39
            
    End Select
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    oForm.Freeze (False)
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i As Long
    On Error GoTo Raise_FormMenuEvent_Error
    Call oForm.Freeze(True)
    
    If (pval.BeforeAction = True) Then
        Select Case pval.MenuUID
            Case "1283":
                If Sbo_Application.MessageBox("ÇöÀç È­¸é³»¿ëÀüÃ¼¸¦ Á¦°Å ÇÏ½Ã°Ú½À´Ï±î? º¹±¸ÇÒ ¼ö ¾ø½À´Ï´Ù.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1284":
            Case "1286":
            Case "1293":
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY650_FormItemEnabled
                Call PH_PY650_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//¹®¼­Ã£±â
                Call PH_PY650_FormItemEnabled
                Call PH_PY650_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//¹®¼­Ãß°¡
                Call PH_PY650_FormItemEnabled
                Call PH_PY650_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY650_FormItemEnabled
            Case "1293" '// Çà»èÁ¦
'                '// [MAT1 ¿ë]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY650B.Size - 1)
'                        If oDS_PH_PY650B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY650B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY650B.Size
'                        Call oDS_PH_PY650B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
                'End If
                Call PH_PY650_AddMatrixRow
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_FormDataEvent_Error
    
    If (BusinessObjectInfo.BeforeAction = True) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select
    End If
    Exit Sub
    
Raise_FormDataEvent_Error:
    
        Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_RightClickEvent_Error
    
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
    End If
    Select Case pval.ItemUID
    Case "Mat01"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUID
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUID
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY650_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY650_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 ¿ë]
    'oMat1.FlushToDataSource
    'oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY650B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY650B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY650B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY650B.Offset = oRow
'            oDS_PH_PY650B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY650B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY650B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY650B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY650B.Offset = oRow - 1
'            oDS_PH_PY650B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY650B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY650B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY650B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY650B.Offset = oRow
'        oDS_PH_PY650B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY650B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY650B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY650B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY650_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY650_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY650_FormClear()
On Error GoTo PH_PY650_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY650'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY650_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY650_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY650_DataValidCheck() As Boolean
On Error GoTo PH_PY650_DataValidCheck_Error
    PH_PY650_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    '//----------------------------------------------------------------------------------
'    '//ÇÊ¼ö Ã¼Å©
'    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY650A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "»ç¿ø¹øÈ£´Â ÇÊ¼öÀÔ´Ï´Ù.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY650_DataValidCheck = False
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix ¸¶Áö¸· Çà »èÁ¦(DB ÀúÀå½Ã)
'    If oDS_PH_PY650B.Size > 1 Then oDS_PH_PY650B.RemoveRecord (oDS_PH_PY650B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY650_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY650_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY650_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY650_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY650_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY650_Validate_Error
    PH_PY650_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY650A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "ÇØ´ç¹®¼­´Â ´Ù¸¥»ç¿ëÀÚ¿¡ ÀÇÇØ Ãë¼ÒµÇ¾ú½À´Ï´Ù. ÀÛ¾÷À» ÁøÇàÇÒ¼ö ¾ø½À´Ï´Ù.", bmt_Short, True
        PH_PY650_Validate = False
        GoTo PH_PY650_Validate_Exit
    End If
'
    If ValidateType = "¼öÁ¤" Then

    ElseIf ValidateType = "Çà»èÁ¦" Then

    ElseIf ValidateType = "Ãë¼Ò" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY650_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY650_Validate_Error:
    PH_PY650_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY650_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY650_Print_Report01()
On Error GoTo PH_PY650_Print_Report01_Error

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Dim CLTCOD          As String
    Dim DocDate         As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("Á¶È¸ Áß‘...", 100, False)
    
     '/ ODBC ¿¬°á Ã¼Å©
    If ConnectODBC = False Then
        GoTo PH_PY650_Print_Report01_Error
    End If
    
    '//ÀÎÀÚ MOVE , Trim ½ÃÅ°±â..
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    DocDate = oForm.Items("DocDate").Specific.VALUE
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY650] ³ëµ¿Á¶ÇÕ°£ºÎÇöÈ²"
    ReportName = "PH_PY650_01.rpt"
    
    '/ Formula ¼ö½ÄÇÊµå
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
    gRpt_Formula(1) = "CLTCOD"
    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
    
    gRpt_Formula(2) = "DocDate"
    gRpt_Formula_Value(2) = Format(DocDate, "####-##-##")
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure ½ÇÇà"
    sQry = "EXEC [PH_PY650_01] '" & CLTCOD & "', '" & DocDate & "'"
    
    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY650_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V", , 1) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : ½ÇÆÐ!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY650_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Ãâ·ÂÇÒ µ¥ÀÌÅÍ°¡ ¾ø½À´Ï´Ù. È®ÀÎÇØ ÁÖ¼¼¿ä.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY650_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub













