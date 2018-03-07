VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY121.cls
'//  Module         : �λ���� > �޿�����
'//  Desc           : ���κ� �򰡰��޾� ���
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY121A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY121B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY121.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY121_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY121"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY121_CreateItems
    Call PH_PY121_EnableMenus
    Call PH_PY121_SetDocument(oFromDocEntry01)
'    Call PH_PY121_FormResize
    
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

Private Function PH_PY121_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY121_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY121A = oForm.DataSources.DBDataSources("@PH_PY121A")   '//���
    Set oDS_PH_PY121B = oForm.DataSources.DBDataSources("@PH_PY121B")   '//����

    Set oMat1 = oForm.Items("Mat1").Specific       '

    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns

    '//----------------------------------------------------------------------------------------------
    '// ������ ����
    '//----------------------------------------------------------------------------------------------
    '//�����
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '��������
    Set oCombo = oForm.Items("JIGTYP").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P126' AND U_UseYN= 'Y' "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JIGTYP").DisplayDesc = True
    
    
    Set oColumn = oMat1.Columns("UseYN")
    oColumn.ValidValues.Add "Y", "Y(���)"
    oColumn.ValidValues.Add "N", "N(������)"
    oColumn.DisplayDesc = True
    
    
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY121_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY121_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function




Private Sub PH_PY121_EnableMenus()
    
    On Error GoTo PH_PY121_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//����
    Call oForm.EnableMenu("1284", False)    '//���
    Call oForm.EnableMenu("1293", True)     '//�����
    
    Exit Sub
    
PH_PY121_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY121_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY121_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY121_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY121_FormItemEnabled
        Call PH_PY121_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY121_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY121_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY121_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY121_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oConditions As SAPbouiCOM.Conditions
    
    On Error GoTo PH_PY121_FormItemEnabled_Error
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YEAR").Enabled = True
        oForm.Items("YMFROM").Enabled = True
        oForm.Items("YMTO").Enabled = True
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        '/ �ͼӳ��
        oForm.Items("YEAR").Specific.VALUE = Format$(Now, "YYYY")
        
        
        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", False)    '//�����߰�
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YEAR").Enabled = True
        oForm.Items("YMFROM").Enabled = True
        oForm.Items("YMTO").Enabled = True
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YEAR").Enabled = False
        oForm.Items("YMFROM").Enabled = False
        oForm.Items("YMTO").Enabled = False
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY121_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY121_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY121_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    End If
                End If
                
                If pval.ItemUID = "Btn_Create" Then
                    Call MSTCOD_Setting
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If pval.ActionSuccess = True Then
                        Call PH_PY121_FormItemEnabled
                        
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True And pval.ItemUID = "Mat1" And pval.ColUID = "MSTCOD" And pval.CharPressed = 9 Then
                If oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE = "" Then
                    oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                Else
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE + "'") = True Then
                        oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat1"
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

            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat1"
                    If pval.Row > 0 Then
                        '//Call oMat1.SelectRow(pval.Row, True, False)

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
                    If pval.ItemUID = "Mat1" And pval.ColUID = "MSTCOD" Then
'                        If oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE <> "" Then
'                            oMat1.Columns("MSTNAM").Cells(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE & "'")
'                            oMat1.Columns("DPTCOD").Cells(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE & "'")
'                            oMat1.Columns("DPTNAM").Cells(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat1.Columns("DPTCOD").Cells(pval.Row).Specific.VALUE & "'", " AND Code = '1'")
'                        End If
                        Call PH_PY121_AddMatrixRow
                        oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY121_FormItemEnabled
                Call PH_PY121_AddMatrixRow

            End If
'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            '�÷��ǿ��� ������ ��� �޸� ����
            '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_PH_PY121A = Nothing
                Set oDS_PH_PY121B = Nothing

                Set oMat1 = Nothing
            End If
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.BeforeAction = False Then
'
'            End If
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
                If Sbo_Application.MessageBox("���� ȭ�鳻����ü�� ���� �Ͻðڽ��ϱ�? ������ �� �����ϴ�.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1284":
            Case "1286":
            Case "1293":
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY121A", "Code")      '//������ ���ѿ� ���� ����� ����
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY121_FormItemEnabled
                Call PH_PY121_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY121_FormItemEnabled
                Call PH_PY121_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY121_FormItemEnabled
                Call PH_PY121_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY121_FormItemEnabled
            Case "1293" '// �����
'                '// [MAT1 ��]
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource

                    While (i <= oDS_PH_PY121B.Size - 1)
                        If oDS_PH_PY121B.GetValue("U_MSTCOD", i) = "" Then
                            oDS_PH_PY121B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY121B.Size
                        Call oDS_PH_PY121B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat1.LoadFromDataSource
                End If
                Call PH_PY121_AddMatrixRow
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Case "Mat1"
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

Sub PH_PY121_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY121_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 ��]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount

    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY121B.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY121B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY121B.InsertRecord (oRow)
            End If
            oDS_PH_PY121B.Offset = oRow
            oDS_PH_PY121B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY121B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY121B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY121B.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY121B.setValue "U_TeamName", oRow, ""
            oDS_PH_PY121B.setValue "U_RspCode", oRow, ""
            oDS_PH_PY121B.setValue "U_RspName", oRow, ""
            oDS_PH_PY121B.setValue "U_Grade", oRow, ""
            oDS_PH_PY121B.setValue "U_Amt", oRow, 0
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY121B.Offset = oRow - 1
            oDS_PH_PY121B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY121B.setValue "U_MSTCOD", oRow - 1, ""
            oDS_PH_PY121B.setValue "U_MSTNAM", oRow - 1, ""
            oDS_PH_PY121B.setValue "U_TeamCode", oRow - 1, ""
            oDS_PH_PY121B.setValue "U_TeamName", oRow - 1, ""
            oDS_PH_PY121B.setValue "U_RspCode", oRow - 1, ""
            oDS_PH_PY121B.setValue "U_RspName", oRow - 1, ""
            oDS_PH_PY121B.setValue "U_Grade", oRow - 1, ""
            oDS_PH_PY121B.setValue "U_Amt", oRow - 1, 0
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY121B.Offset = oRow
        oDS_PH_PY121B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY121B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY121B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY121B.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY121B.setValue "U_TeamName", oRow, ""
            oDS_PH_PY121B.setValue "U_RspCode", oRow, ""
            oDS_PH_PY121B.setValue "U_RspName", oRow, ""
            oDS_PH_PY121B.setValue "U_Grade", oRow, ""
            oDS_PH_PY121B.setValue "U_Amt", oRow, 0
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY121_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY121_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY121_FormClear()
On Error GoTo PH_PY121_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY121'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY121_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY121_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY121_DataValidCheck() As Boolean
On Error GoTo PH_PY121_DataValidCheck_Error
    PH_PY121_DataValidCheck = False
    Dim i           As Long
    Dim k           As Long
    Dim SCode       As String
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//�ʼ�üũ
    '//----------------------------------------------------------------------------------
    '// ���
    If Trim(oDS_PH_PY121A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "������� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY121A.GetValue("U_YEAR", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "�������� �⵵�� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("YEAR").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY121A.GetValue("U_YMFROM", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "�򰡰��޾� ���� ���ۿ� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("YMFROM").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY121A.GetValue("U_YMTO", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "�򰡰��޾� ����������� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("YMTO").CLICK ct_Regular
        Exit Function
    End If
    
    oMat1.FlushToDataSource
    '// ���� ---------------------------
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            '//����
            If oMat1.Columns("MSTCOD").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "����� �ʼ��Դϴ�.", bmt_Short, True
                oMat1.Columns("MSTCOD").Cells(i).CLICK ct_Regular
                Exit Function
            End If
           
            '//�ߺ�üũ
            For k = i + 1 To oMat1.VisualRowCount - 2
                If Trim$(oDS_PH_PY121B.GetValue("U_MSTCOD", i)) = Trim$(oDS_PH_PY121B.GetValue("U_MSTCOD", k)) Then
                    Sbo_Application.SetStatusBarMessage "������ �ߺ��ԷµǾ����ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
                    oMat1.Columns("MSTCOD").Cells(k + 1).CLICK ct_Regular
                    Exit Function
                End If
            Next k
        Next
    Else
        Sbo_Application.SetStatusBarMessage "���� �����Ͱ� �����ϴ�.", bmt_Short, True
        PH_PY121_DataValidCheck = False
        Exit Function
    End If
    
    SCode = Trim$(oDS_PH_PY121A.GetValue("U_YEAR", 0)) & Trim$(oDS_PH_PY121A.GetValue("U_CLTCOD", 0)) & Trim$(oDS_PH_PY121A.GetValue("U_JIGTYP", 0))
    
    oDS_PH_PY121A.setValue "Code", 0, SCode
    oDS_PH_PY121A.setValue "Name", 0, SCode
    
    oMat1.FlushToDataSource
    '// Matrix ������ �� ����(DB �����)
    If oDS_PH_PY121B.Size > 1 Then oDS_PH_PY121B.RemoveRecord (oDS_PH_PY121B.Size - 1)
    oMat1.LoadFromDataSource

    '//������ �ߺ� üũ
    If MDC_SetMod.Get_ReData("Code", "Code", "[@PH_PY121A]", "'" & "sCode" & "'") <> "" Then
        Sbo_Application.SetStatusBarMessage "������ �����Ͱ� �����մϴ�.", bmt_Short, True
        PH_PY121_DataValidCheck = False
        Exit Function
    End If
    
    PH_PY121_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY121_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY121_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY121_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY121_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY121_Validate_Error
    PH_PY121_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY121A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY121_Validate = False
        GoTo PH_PY121_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY121_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY121_Validate_Error:
    PH_PY121_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY121_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub MSTCOD_Setting()
On Error GoTo Error_Message
    Dim i           As Long
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim iRow        As Long
    Dim MatRowCnt   As Long:    Dim oRow  As Long
    Dim ErrNum      As Integer
    
    Dim CLTCOD      As String   '//�����
    Dim YEAR      As String   '//�ͼӳ��
    Dim JIGTYP      As String '//��������
    
    
    Dim V_StatusCnt&, oProValue&, TOTCNT&, tRow&  '//progbar
    
    Dim CheckValue  As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    oForm.Freeze (True)
    oMat1.Clear
    oDS_PH_PY121B.Clear
    
    CLTCOD = Trim$(oDS_PH_PY121A.GetValue("U_CLTCOD", 0))
    YEAR = Trim$(oDS_PH_PY121A.GetValue("U_YEAR", 0))
    JIGTYP = Trim$(oDS_PH_PY121A.GetValue("U_JIGTYP", 0))
    If YEAR = "" Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    If JIGTYP = "" Then
        ErrNum = 4
        GoTo Error_Message
    End If
    
    sQry = "EXEC PH_PY121_01 '" & CLTCOD & "','" & YEAR & "','" & JIGTYP & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    
    Set oProgBar = Sbo_Application.StatusBar.CreateProgressBar("������ �д���...!", 50, False)
    
    '�ִ밪 ���ϱ� ///////////////////////////////////////
    TOTCNT = oRecordSet.RecordCount
   
    V_StatusCnt = Round(TOTCNT / 50, 0)
    oProValue = 1
    tRow = 1
    '/////////////////////////////////////////////////////
    
    
'    oDS_PH_PY121B.RemoveRecord (oDS_PH_PY121B.Size - 1)
    
    MatRowCnt = oMat1.VisualRowCount - 2
'    oRow = oMat1.VisualRowCount - 1
    Do Until oRecordSet.EOF
        
        If oRow + 1 > oDS_PH_PY121B.Size Then
            oDS_PH_PY121B.InsertRecord (oRow)
        End If
        
'        oDS_PH_PY121B.InsertRecord oRow
'        oMat1.AddRow
        oDS_PH_PY121B.Offset = oRow
        oDS_PH_PY121B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY121B.setValue "U_MSTCOD", oRow, Trim(oRecordSet.Fields("MSTCOD").VALUE)    '�����ȣ
        oDS_PH_PY121B.setValue "U_MSTNAM", oRow, Trim(oRecordSet.Fields("MSTNAM").VALUE)    '����
        oDS_PH_PY121B.setValue "U_TeamCode", oRow, Trim(oRecordSet.Fields("TeamCode").VALUE)    '�μ�
        oDS_PH_PY121B.setValue "U_TeamName", oRow, Trim(oRecordSet.Fields("TeamName").VALUE)    '�μ���
        oDS_PH_PY121B.setValue "U_RspCode", oRow, Trim(oRecordSet.Fields("RspCode").VALUE)    '���
        oDS_PH_PY121B.setValue "U_RspName", oRow, Trim(oRecordSet.Fields("RspName").VALUE) '����
        oDS_PH_PY121B.setValue "U_Grade", oRow, Trim(oRecordSet.Fields("Grade").VALUE)
        oDS_PH_PY121B.setValue "U_Amt", oRow, Trim(oRecordSet.Fields("Amt").VALUE)
        oDS_PH_PY121B.setValue "U_UseYN", oRow, Trim(oRecordSet.Fields("UseYN").VALUE)
        oRow = oRow + 1
        
        
        tRow = tRow + 1
        If (TOTCNT > 50 And tRow = oProValue * V_StatusCnt) Or TOTCNT <= 50 Then
            oProgBar.Text = tRow & "/ " & TOTCNT & " �� ó����...!"
            oProValue = oProValue + 1
            oProgBar.VALUE = oProValue
        End If
        
        oRecordSet.MoveNext
    Loop
    
    Sbo_Application.StatusBar.SetText "�۾��� �Ϸ�Ǿ����ϴ�.", bmt_Short, smt_Success
     
    If oRow = MatRowCnt + 1 Then
        ErrNum = 3
        GoTo Error_Message
    End If
'    oForm.Freeze (False)
    
   
    
    oMat1.LoadFromDataSource
    Call PH_PY121_AddMatrixRow
    
    oProgBar.Stop
    Set oProgBar = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "�򰡳⵵�� �ʼ��Դϴ�. �Է��Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "������ ����ڰ� �����ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "�̹� ��� ����ڰ� �����Ǿ� �־� �߰������� ����ڰ� �����ϴ�.", bmt_Short, smt_Success
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "���������� �ʼ��Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "MSTCOD_Setting Error : " & Err.Description, bmt_Short, smt_Error
    End If
    
'    oForm.Freeze (False)
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    Set oRecordSet = Nothing
End Sub

