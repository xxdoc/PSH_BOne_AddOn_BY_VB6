VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY302.cls
'//  Module         : �λ���� > ��Ÿ
'//  Desc           : ���ڱ����޿Ϸ�ó��
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY302A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY302B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY302.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY302_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY302"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY302_CreateItems
    Call PH_PY302_EnableMenus
    Call PH_PY302_SetDocument(oFromDocEntry01)
'    Call PH_PY302_FormResize
    
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

Private Function PH_PY302_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY302_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY302A = oForm.DataSources.DBDataSources("@PH_PY302A")
    Set oDS_PH_PY302B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    Set oMat1 = oForm.Items("Mat01").Specific

    oMat1.SelectionMode = ms_Auto
    oMat1.AutoResizeColumns


    '�����_S
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 5)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    '�����_E
    
    '�⵵_S
    Call oForm.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 4)
    Call oForm.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")
    '�⵵_E
        
    '�б�_S
    Call oForm.DataSources.UserDataSources.Add("Quarter", dt_SHORT_TEXT, 5)
    Call oForm.Items("Quarter").Specific.DataBind.SetBound(True, "", "Quarter")
    '�б�_E

    '���бݰ�_S
    Call oForm.DataSources.UserDataSources.Add("FeeTot", dt_SUM, 50)
    Call oForm.Items("FeeTot").Specific.DataBind.SetBound(True, "", "FeeTot")
    '���бݰ�_E

    '��ϱݰ�_S
    Call oForm.DataSources.UserDataSources.Add("TuiTot", dt_SUM, 50)
    Call oForm.Items("TuiTot").Specific.DataBind.SetBound(True, "", "TuiTot")
    '��ϱݰ�_E

    '�Ѱ�_S
    Call oForm.DataSources.UserDataSources.Add("Total", dt_SUM, 50)
    Call oForm.Items("Total").Specific.DataBind.SetBound(True, "", "Total")
    '�Ѱ�_E

    '//----------------------------------------------------------------------------------------------
    '// �⺻����
    '//----------------------------------------------------------------------------------------------
    
    '�����
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '�б�
    Set oCombo = oForm.Items("Quarter").Specific
    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "01", "1/4 Ȥ�� 1�б�"
    oCombo.ValidValues.Add "02", "2/4"
    oCombo.ValidValues.Add "03", "3/4 Ȥ�� 2�б�"
    oCombo.ValidValues.Add "04", "4/4"
    oCombo.Select 0, psk_Index
    oForm.Items("Quarter").DisplayDesc = True
    
    '��Ʈ����-ȸ��
    Set oColumn = oMat1.Columns("Count")
    oColumn.ValidValues.Add "", ""
    oColumn.ValidValues.Add "01", "1��"
    oColumn.ValidValues.Add "02", "2��"
    oColumn.DisplayDesc = True



    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY302_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY302_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY302_EnableMenus()
    
    On Error GoTo PH_PY302_EnableMenus_Error
    
'    Call oForm.EnableMenu("1283", False)         '// ����
'    Call oForm.EnableMenu("1287", True)          '// ����
''    Call oForm.EnableMenu("1286", True)         '// �ݱ�
'    Call oForm.EnableMenu("1284", True)         '// ���
'    Call oForm.EnableMenu("1293", True)         '// �����
    
    Exit Sub
    
PH_PY302_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY302_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY302_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY302_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY302_FormItemEnabled
'        Call PH_PY302_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY302_FormItemEnabled
        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY302_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY302_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY302_FormItemEnabled()
    On Error GoTo PH_PY302_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim CLTCOD As String
        
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        '�� DocEntry ����
'        Call PH_PY302_FormClear
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '�⵵ ����
'        Call oDS_PH_PY302A.setValue("U_StdYear", 0, Format(Date, "YYYY"))
        oForm.Items("StdYear").Specific.Value = Format(Date, "YYYY")

        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", False)    '//�����߰�
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY302_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY302_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry As String
    Dim i As Long
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Dim loopCount As Integer 'For Loop �� (VALIDATE Event���� ���)
    Dim GovID1 As String '�ֹε�Ϲ�ȣ ���ڸ�(VALIDATE Event���� ���)
    Dim GovID2 As String '�ֹε�Ϲ�ȣ ���ڸ�(VALIDATE Event���� ���)
    Dim GovID As String '�ֹε�Ϲ�ȣ ��ü(VALIDATE Event���� ���)
    Dim Sex As String '����(VALIDATE Event���� ���)
    Dim PayCnt As Integer '����Ƚ��(VALIDATE Event���� ���)
    Dim FeeTot As Double '���бݰ�(VALIDATE Event���� ���)
    Dim TuiTot As Double '��ϱݰ�(VALIDATE Event���� ���)
    Dim Total As Double '�Ѱ�(VALIDATE Event���� ���)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If PH_PY302_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'
'                        '//�ؾ����� �۾�
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If PH_PY302_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'                        '//�ؾ����� �۾�
'
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                    End If
'                End If
                If pval.ItemUid = "BtnSearch" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY302_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'
'                        '//�ؾ����� �۾�
                        Call PH_PY302_MTX01
                        
                    End If
                    
                ElseIf pval.ItemUid = "BtnModify" Then
                
                    If PH_PY302_UpdatePayYN = True Then
                        Call MDC_Com.MDC_GF_Message("���޿ϷῩ�ΰ� �����Ǿ����ϴ�.", "S")
                    Else
                        BubbleEvent = False
                        Exit Sub
                    End If
                
                End If
                
            ElseIf pval.BeforeAction = False Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY302_FormItemEnabled
'                            Call PH_PY302_AddMatrixRow
'                        End If
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY302_FormItemEnabled
'                            Call PH_PY302_AddMatrixRow
'                        End If
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY302_FormItemEnabled
'                        End If
'                    End If
'                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                If pval.ItemUid = "Mat01" Then
        
                    If pval.ColUID = "Name" And pval.CharPressed = "9" Then
                    
                        If oMat1.Columns.Item("Name").Cells(pval.Row).Specific.Value = "" Then
                            Call Sbo_Application.ActivateMenuItem("7425")
                            BubbleEvent = False
                        End If
                    
                    End If
                
                ElseIf pval.ItemUid = "CntcCode" And pval.CharPressed = "9" Then
                
                    If oForm.Items("CntcCode").Specific.Value = "" Then
                        Call Sbo_Application.ActivateMenuItem("7425")
                        BubbleEvent = False
                    End If
            
                End If

            ElseIf pval.Before_Action = False Then

            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUid
            Case "Mat01"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUid
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUid
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
'                    Call PH_PY302_AddMatrixRow
                    Call oMat1.AutoResizeColumns
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUid
                Case "Mat01"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
                End Select
                
                Select Case pval.ItemUid
                Case "Mat01"
                    If pval.Row > 0 Then
                        oLastItemUID = pval.ItemUid
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case Else
                    oLastItemUID = pval.ItemUid
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
            
                If pval.ItemChanged = True Then

                End If
       
            ElseIf pval.BeforeAction = False Then
            
                If pval.ItemChanged = True Then
                    
                    Select Case pval.ItemUid
                        
                    End Select

                End If
                
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY302_FormItemEnabled
                Call PH_PY302_AddMatrixRow
                Call oMat1.AutoResizeColumns
                
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
                Set oDS_PH_PY302A = Nothing
                Set oDS_PH_PY302B = Nothing
                
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

                Call oMat1.AutoResizeColumns

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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY302A", "Code")
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
    
    Dim loopCount As Integer
    Dim FeeTot As Double
    Dim TuiTot As Double
    Dim Total As Double
    
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
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY302_FormItemEnabled
                Call PH_PY302_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY302_FormItemEnabled
                Call PH_PY302_AddMatrixRow
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY302_FormItemEnabled
                Call PH_PY302_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY302_FormItemEnabled
            Case "1293" '// �����

                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY302B.Size - 1)
                        If oDS_PH_PY302B.GetValue("U_LineNum", i) = "" Then
                            oDS_PH_PY302B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY302B.Size
                        Call oDS_PH_PY302B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY302_AddMatrixRow
                
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
    Select Case pval.ItemUid
    Case "Mat01"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUid
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUid
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY302_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY302_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY302B.GetValue("U_Name", oRow - 1)) <> "" Then
            If oDS_PH_PY302B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY302B.InsertRecord (oRow)
            End If
            oDS_PH_PY302B.Offset = oRow
            oDS_PH_PY302B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY302B.setValue "U_ColReg01", oRow, ""
            oDS_PH_PY302B.setValue "U_ColSum01", oRow, ""
            oDS_PH_PY302B.setValue "U_ColSum02", oRow, ""
            oDS_PH_PY302B.setValue "U_ColSum03", oRow, ""
            oDS_PH_PY302B.setValue "U_ColReg02", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY302B.Offset = oRow - 1
            oDS_PH_PY302B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY302B.setValue "U_ColReg01", oRow - 1, ""
            oDS_PH_PY302B.setValue "U_ColSum01", oRow - 1, ""
            oDS_PH_PY302B.setValue "U_ColSum02", oRow - 1, ""
            oDS_PH_PY302B.setValue "U_ColSum03", oRow - 1, ""
            oDS_PH_PY302B.setValue "U_ColReg02", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY302B.Offset = oRow
        oDS_PH_PY302B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY302B.setValue "U_ColReg01", oRow, ""
        oDS_PH_PY302B.setValue "U_ColSum01", oRow, ""
        oDS_PH_PY302B.setValue "U_ColSum02", oRow, ""
        oDS_PH_PY302B.setValue "U_ColSum03", oRow, ""
        oDS_PH_PY302B.setValue "U_ColReg02", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY302_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY302_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY302_FormClear()
On Error GoTo PH_PY302_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY302'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY302_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY302_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY302_DataValidCheck() As Boolean
On Error GoTo PH_PY302_DataValidCheck_Error
    PH_PY302_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '�����
    If Trim(oForm.Items("CLTCOD").Specific.Value) = "" Then
        Sbo_Application.SetStatusBarMessage "������� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY302_DataValidCheck = False
        Exit Function
    End If
    
    '�⵵
    If Trim(oForm.Items("StdYear").Specific.Value) = "" Then
        Sbo_Application.SetStatusBarMessage "�⵵�� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("StdYear").CLICK ct_Regular
        PH_PY302_DataValidCheck = False
        Exit Function
    End If
    
    '�б�
    If Trim(oForm.Items("Quarter").Specific.Value) = "" Then
        Sbo_Application.SetStatusBarMessage "�б�� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("Quarter").CLICK ct_Regular
        PH_PY302_DataValidCheck = False
        Exit Function
    End If
    
    
'    '����
'    If oMat1.VisualRowCount > 1 Then
'        For i = 1 To oMat1.VisualRowCount - 1
'
'            '�б�
'            If oMat1.Columns("SchCls").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "�б��� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("SchCls").Cells(i).CLICK ct_Regular
'                PH_PY302_DataValidCheck = False
'                Exit Function
'            End If
'
'            '�б���
'            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "�б����� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
'                PH_PY302_DataValidCheck = False
'                Exit Function
'            End If
'
'            '�г�
'            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "�г��� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
'                PH_PY302_DataValidCheck = False
'                Exit Function
'            End If
'
'            'ȸ��
'            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "ȸ���� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
'                PH_PY302_DataValidCheck = False
'                Exit Function
'            End If
'
'        Next
'    Else
'        Sbo_Application.SetStatusBarMessage "���� �����Ͱ� �����ϴ�.", bmt_Short, True
'        PH_PY302_DataValidCheck = False
'        Exit Function
'    End If
    
        oMat1.FlushToDataSource
    '// Matrix ������ �� ����(DB �����)
    If oDS_PH_PY302B.Size > 1 Then oDS_PH_PY302B.RemoveRecord (oDS_PH_PY302B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY302_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY302_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY302_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY302_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY302_MTX01()

    '//��Ʈ������ ������ �ε�
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
'    Dim Param04 As String
    
    Dim FeeTot As Double '���бݰ�
    Dim TuiTot As Double '��ϱݰ�
    Dim Total As Double '�Ѱ�
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY302_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("CLTCOD").Specific.Value
    Param02 = oForm.Items("StdYear").Specific.Value
    Param03 = oForm.Items("Quarter").Specific.Value
'    Param04 = oForm.Items("Param01").Specific.Value
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet.RecordCount, False)
    
    sQry = "EXEC PH_PY302_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call oMat1.Clear
        GoTo PH_PY302_MTX01_Exit
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY302B.InsertRecord (i)
        End If
        oDS_PH_PY302B.Offset = i
        Call oDS_PH_PY302B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY302B.setValue("U_ColReg01", i, oRecordSet.Fields("Count").Value)
        Call oDS_PH_PY302B.setValue("U_ColSum01", i, oRecordSet.Fields("EntFee").Value)
        Call oDS_PH_PY302B.setValue("U_ColSum02", i, oRecordSet.Fields("Tuition").Value)
        Call oDS_PH_PY302B.setValue("U_ColSum03", i, oRecordSet.Fields("Total").Value)
        Call oDS_PH_PY302B.setValue("U_ColReg02", i, oRecordSet.Fields("PayYN").Value)
        
        FeeTot = FeeTot + oRecordSet.Fields("EntFee").Value
        TuiTot = TuiTot + oRecordSet.Fields("Tuition").Value
        
        oRecordSet.MoveNext
        ProgressBar01.Value = ProgressBar01.Value + 1
        ProgressBar01.Text = ProgressBar01.Value & "/" & oRecordSet.RecordCount & "�� ��ȸ��...!"
        
    Next
    
    Total = FeeTot + TuiTot
    
    oForm.Items("FeeTot").Specific.Value = FeeTot
    oForm.Items("TuiTot").Specific.Value = TuiTot
    oForm.Items("Total").Specific.Value = Total
    
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY302_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "W")
    Exit Sub
PH_PY302_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY302_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY302_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY302_Validate_Error
    PH_PY302_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY302A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY302_Validate = False
        GoTo PH_PY302_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY302_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY302_Validate_Error:
    PH_PY302_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY302_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY302_UpdatePayYN() As Boolean
On Error GoTo PH_PY302_UpdatePayYN_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim CLTCOD As String '�����
    Dim StdYear As String '�⵵
    Dim Quarter As String '�б�
    Dim Count As String 'ȸ��
    Dim PayYN As String '���޿ϷῩ��

    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = oForm.Items("CLTCOD").Specific.Value
    StdYear = oForm.Items("StdYear").Specific.Value
    Quarter = oForm.Items("Quarter").Specific.Value
'    Count = oForm.Items("Count").Specific.Value
        
    oMat1.FlushToDataSource
    For loopCount = 0 To oMat1.RowCount - 1
        
        Count = Trim(oDS_PH_PY302B.GetValue("U_ColReg01", loopCount))
        PayYN = Trim(oDS_PH_PY302B.GetValue("U_ColReg02", loopCount))
    
        sQry = "EXEC PH_PY302_02 '" & CLTCOD & "','" & StdYear & "','" & Quarter & "','" & Count & "','" & PayYN & "'"
        Call oRecordSet.DoQuery(sQry)
    
    Next
    
    PH_PY302_UpdatePayYN = True
    
    Exit Function
PH_PY302_UpdatePayYN_Error:
    
    PH_PY302_UpdatePayYN = False
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY302_UpdatePayYN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function








