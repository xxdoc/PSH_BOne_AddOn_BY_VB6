VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY102.cls
'//  Module         : �޿����� > �޿�����
'//  Desc           : �����׸���
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY102A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY102B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY102.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY102_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY102"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY102_CreateItems
    Call PH_PY102_EnableMenus
    Call PH_PY102_SetDocument(oFromDocEntry01)
'    Call PH_PY102_FormResize
    
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

Private Function PH_PY102_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY102_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY102A = oForm.DataSources.DBDataSources("@PH_PY102A")
    Set oDS_PH_PY102B = oForm.DataSources.DBDataSources("@PH_PY102B")

    
    Set oMat1 = oForm.Items("Mat1").Specific       '//@PH_PY102B

    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns


    '// ��� ------------------------------------------------------------------------------
    '// �����
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// ���� ------------------------------------------------------------------------------
    
    '// �����ڵ�
    Set oColumn = oMat1.Columns("CSUCOD")
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P134' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = False
    
    '// ���籸��
    Set oColumn = oMat1.Columns("CSUGBN")
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P141' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True
    
    '// ��������
    Set oColumn = oMat1.Columns("GWATYP")
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P142' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True
    
    '// �����׸�
    Set oColumn = oMat1.Columns("FIXGBN")
    oColumn.ValidValues.Add "Y", "����"
    oColumn.ValidValues.Add "N", "���"
    oColumn.ValidValues.Add "V", "����"
    oColumn.DisplayDesc = True
    
    '// ����ó��
    Set oColumn = oMat1.Columns("ROUNDT")
    oColumn.ValidValues.Add "R", "�ݿø�"
    oColumn.ValidValues.Add "F", "����"
    oColumn.ValidValues.Add "C", "����"
    oColumn.DisplayDesc = True
    
    '// ����
    Set oColumn = oMat1.Columns("LENGTH")
    oColumn.ValidValues.Add "1", "��"
    oColumn.ValidValues.Add "10", "�ʿ�"
    oColumn.ValidValues.Add "100", "���"
    oColumn.ValidValues.Add "1000", "õ��"
    oColumn.DisplayDesc = True

    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY102_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY102_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY102_EnableMenus()
    
    On Error GoTo PH_PY102_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//����
    Call oForm.EnableMenu("1284", False)    '//���
    Call oForm.EnableMenu("1293", True)     '//�����
    
    Exit Sub
    
PH_PY102_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY102_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY102_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY102_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY102_FormItemEnabled
        Call PH_PY102_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY102_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY102_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY102_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY102_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY102_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", False)    '//�����߰�
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YM").Enabled = False
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY102_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY102_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY102_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        
                        '//�ؾ����� �۾�
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY102_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        '//�ؾ����� �۾�
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY102_FormItemEnabled
                            Call PH_PY102_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY102_FormItemEnabled
                            Call PH_PY102_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY102_FormItemEnabled
                        End If
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

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
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat1" And pval.ColUID = "CSUCOD" Then
                        oMat1.Columns("CSUNAM").Cells(pval.Row).Specific.Value = MDC_SetMod.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat1.Columns("CSUCOD").Cells(pval.Row).Specific.Value & "'", " AND Code = 'P134'")
                        Call PH_PY102_AddMatrixRow
                        oMat1.Columns("LINSEQ").Cells(pval.Row).CLICK ct_Regular
                        oMat1.AutoResizeColumns
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat1"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
                End Select
                
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

                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY102_FormItemEnabled
                Call PH_PY102_AddMatrixRow
                
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
                Set oDS_PH_PY102A = Nothing
                Set oDS_PH_PY102B = Nothing
                
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
'                If pval.ItemUID = "Code" Then
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY102A", "Code")
'                End If
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY102A", "Code")      '//������ ���ѿ� ���� ����� ����
            
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY102_FormItemEnabled
                Call PH_PY102_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY102_FormItemEnabled
                Call PH_PY102_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY102_FormItemEnabled
                Call PH_PY102_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY102_FormItemEnabled
            Case "1293" '// �����
                '// [MAT1] �޿� ����
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY102B.Size - 1)
                        If oDS_PH_PY102B.GetValue("U_CSUCOD", i) = "" Then
                            oDS_PH_PY102B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY102B.Size
                        Call oDS_PH_PY102B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY102_AddMatrixRow
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

Sub PH_PY102_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY102_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY102B.GetValue("U_CSUCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY102B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY102B.InsertRecord (oRow)
            End If
            oDS_PH_PY102B.Offset = oRow
            oDS_PH_PY102B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY102B.setValue "U_CSUCOD", oRow, ""
            oDS_PH_PY102B.setValue "U_CSUNAM", oRow, ""
            oDS_PH_PY102B.setValue "U_LINSEQ", oRow, ""
            oDS_PH_PY102B.setValue "U_CSUGBN", oRow, ""
            oDS_PH_PY102B.setValue "U_GWATYP", oRow, ""
            oDS_PH_PY102B.setValue "U_BTXCOD", oRow, ""
            oDS_PH_PY102B.setValue "U_KUMAMT", oRow, 0
            oDS_PH_PY102B.setValue "U_FIXGBN", oRow, "N"
            oDS_PH_PY102B.setValue "U_INSLIN", oRow, ""
            oDS_PH_PY102B.setValue "U_HOBUSE", oRow, "N"
            oDS_PH_PY102B.setValue "U_BNSUSE", oRow, "N"
            oDS_PH_PY102B.setValue "U_CSUCHK", oRow, "N"
            oDS_PH_PY102B.setValue "U_MONPAY", oRow, "N"
            oDS_PH_PY102B.setValue "U_RETCHK", oRow, "N"
            oDS_PH_PY102B.setValue "U_GBHGBN", oRow, "N"
            oDS_PH_PY102B.setValue "U_ROUNDT", oRow, ""
            oDS_PH_PY102B.setValue "U_LENGTH", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY102B.Offset = oRow - 1
            oDS_PH_PY102B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY102B.setValue "U_CSUCOD", oRow - 1, ""
            oDS_PH_PY102B.setValue "U_CSUNAM", oRow - 1, ""
            oDS_PH_PY102B.setValue "U_LINSEQ", oRow - 1, ""
            oDS_PH_PY102B.setValue "U_CSUGBN", oRow - 1, ""
            oDS_PH_PY102B.setValue "U_GWATYP", oRow - 1, ""
            oDS_PH_PY102B.setValue "U_BTXCOD", oRow - 1, ""
            oDS_PH_PY102B.setValue "U_KUMAMT", oRow - 1, 0
            oDS_PH_PY102B.setValue "U_FIXGBN", oRow - 1, "N"
            oDS_PH_PY102B.setValue "U_INSLIN", oRow - 1, ""
            oDS_PH_PY102B.setValue "U_HOBUSE", oRow - 1, "N"
            oDS_PH_PY102B.setValue "U_BNSUSE", oRow - 1, "N"
            oDS_PH_PY102B.setValue "U_CSUCHK", oRow - 1, "N"
            oDS_PH_PY102B.setValue "U_MONPAY", oRow - 1, "N"
            oDS_PH_PY102B.setValue "U_RETCHK", oRow - 1, "N"
            oDS_PH_PY102B.setValue "U_GBHGBN", oRow - 1, "N"
            oDS_PH_PY102B.setValue "U_ROUNDT", oRow - 1, ""
            oDS_PH_PY102B.setValue "U_LENGTH", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY102B.Offset = oRow
        oDS_PH_PY102B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY102B.setValue "U_CSUCOD", oRow, ""
            oDS_PH_PY102B.setValue "U_CSUNAM", oRow, ""
            oDS_PH_PY102B.setValue "U_LINSEQ", oRow, ""
            oDS_PH_PY102B.setValue "U_CSUGBN", oRow, ""
            oDS_PH_PY102B.setValue "U_GWATYP", oRow, ""
            oDS_PH_PY102B.setValue "U_BTXCOD", oRow, ""
            oDS_PH_PY102B.setValue "U_KUMAMT", oRow, 0
            oDS_PH_PY102B.setValue "U_FIXGBN", oRow, "N"
            oDS_PH_PY102B.setValue "U_INSLIN", oRow, ""
            oDS_PH_PY102B.setValue "U_HOBUSE", oRow, "N"
            oDS_PH_PY102B.setValue "U_BNSUSE", oRow, "N"
            oDS_PH_PY102B.setValue "U_CSUCHK", oRow, "N"
            oDS_PH_PY102B.setValue "U_MONPAY", oRow, "N"
            oDS_PH_PY102B.setValue "U_RETCHK", oRow, "N"
            oDS_PH_PY102B.setValue "U_GBHGBN", oRow, "N"
            oDS_PH_PY102B.setValue "U_ROUNDT", oRow, ""
            oDS_PH_PY102B.setValue "U_LENGTH", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY102_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY102_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY102_FormClear()
On Error GoTo PH_PY102_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY102'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY102_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY102_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY102_DataValidCheck() As Boolean
On Error GoTo PH_PY102_DataValidCheck_Error
    PH_PY102_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '//���
    If Trim(oDS_PH_PY102A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "������� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY102_DataValidCheck = False
        Exit Function
    End If
    
    If Trim(oDS_PH_PY102A.GetValue("U_YM", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "������ۿ��� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        PH_PY102_DataValidCheck = False
        Exit Function
    End If
    
    '// Code & Name ����
    Call oDS_PH_PY102A.setValue("Code", 0, Trim(oDS_PH_PY102A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY102A.GetValue("U_YM", 0)))
    Call oDS_PH_PY102A.setValue("NAME", 0, Trim(oDS_PH_PY102A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY102A.GetValue("U_YM", 0)))
    
    '// ���� ---------------------------
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            '//����
            If oMat1.Columns("CSUCOD").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "�ڵ�� �ʼ��Դϴ�.", bmt_Short, True
                oMat1.Columns("CSUCOD").Cells(i).CLICK ct_Regular
                PH_PY102_DataValidCheck = False
                Exit Function
            End If
            '//�ٷ���
            If oMat1.Columns("CSUNAM").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "������� �ʼ��Դϴ�.", bmt_Short, True
                oMat1.Columns("CSUNAM").Cells(i).CLICK ct_Regular
                PH_PY102_DataValidCheck = False
                Exit Function
            End If
        Next
    Else
        Sbo_Application.SetStatusBarMessage "���� �����Ͱ� �����ϴ�.", bmt_Short, True
        PH_PY102_DataValidCheck = False
        Exit Function
    End If
    
        oMat1.FlushToDataSource
    '// Matrix ������ �� ����(DB �����)
    If oDS_PH_PY102B.Size > 1 Then oDS_PH_PY102B.RemoveRecord (oDS_PH_PY102B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY102_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY102_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY102_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY102_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY102_MTX01()

    '//��Ʈ������ ������ �ε�
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY102_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("Param01").Specific.Value
    Param02 = oForm.Items("Param01").Specific.Value
    Param03 = oForm.Items("Param01").Specific.Value
    Param04 = oForm.Items("Param01").Specific.Value
    
    sQry = "SELECT 10"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "E")
        GoTo PH_PY102_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY102B.InsertRecord (i)
        End If
        oDS_PH_PY102B.Offset = i
        Call oDS_PH_PY102B.setValue("U_COL01", i, oRecordSet.Fields(0).Value)
        Call oDS_PH_PY102B.setValue("U_COL02", i, oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
        ProgressBar01.Value = ProgressBar01.Value + 1
        ProgressBar01.Text = ProgressBar01.Value & "/" & oRecordSet.RecordCount & "�� ��ȸ��...!"
    Next
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY102_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY102_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY102_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY102_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY102_Validate_Error
    PH_PY102_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY102A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY102_Validate = False
        GoTo PH_PY102_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY102_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY102_Validate_Error:
    PH_PY102_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY102_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function







