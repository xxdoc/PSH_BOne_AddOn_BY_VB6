VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY560"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY560.cls
'//  Module         : PH
'//  Desc           : �������Ȳ
'//  �ۼ���         : NGY
'//  DATE           : 2012.12.03
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// �׸��� ����
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PY560     As SAPbouiCOM.DataTable
'
''// ��Ʈ���� ����
'Public oMat1 As SAPbouiCOM.Matrix
'Private oDS_PH_PY560A As SAPbouiCOM.DBDataSource
'Private oDS_PH_PY560B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY560.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY560_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY560"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY560_CreateItems
    Call PH_PY560_EnableMenus
    Call PH_PY560_SetDocument(oFromDocEntry01)
'    Call PH_PY560_FormResize
    
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

Private Function PH_PY560_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY560_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// ������ ����
    '//----------------------------------------------------------------------------------------------
    
    '//�����
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
    Call CLTCOD_Select(oForm, "CLTCOD")
    
    '//��������
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
    oForm.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    oForm.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
    
    '//�μ�
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("TeamCode").Specific
    oCombo.DataBind.SetBound True, "", "TeamCode"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("TeamCode").DisplayDesc = True
    
    
    '//�ְ�
    Set oCombo = oForm.Items("Bus1").Specific
    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "����", "����"
    oCombo.ValidValues.Add "�̿���", "�̿���"
    oCombo.Select 0, psk_Index
    
    '//����
    Set oCombo = oForm.Items("Bus2").Specific
    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "����", "����"
    oCombo.ValidValues.Add "�̿���", "�̿���"
    oCombo.Select 0, psk_Index
    
    '//�߰�
    Set oCombo = oForm.Items("Bus3").Specific
    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "����(05:30)", "����(05:30)"
    oCombo.ValidValues.Add "����(08:30)", "����(08:30)"
    oCombo.ValidValues.Add "�̿���", "�̿���"
    oCombo.Select 0, psk_Index
     
    
    '//üũ�ڽ�1
    Call oForm.DataSources.UserDataSources.Add("Chk1", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("Chk1").Specific
    oCheck.DataBind.SetBound True, "", "Chk1"

    '/üũ�ڽ�2
    Call oForm.DataSources.UserDataSources.Add("Chk2", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("Chk2").Specific
    oCheck.DataBind.SetBound True, "", "Chk2"
    
    'G5 ����
    Call oForm.DataSources.UserDataSources.Add("G5_YN", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("G5_YN").Specific
    oCheck.DataBind.SetBound True, "", "G5_YN"
    oForm.Items("G5_YN").Specific.Checked = True
    
    '�������ϼ� ��ȸ��(2014.09.11 �۸��� �߰�)
    Call oForm.DataSources.UserDataSources.Add("NoAccdnt", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("NoAccdnt").Specific
    oCheck.DataBind.SetBound True, "", "NoAccdnt"
    oForm.Items("NoAccdnt").Specific.Checked = False

'    '//��Ʈ�����÷�
'    Set oColumn = oMat1.Columns("FILD01")
'    oColumn.Editable = True


 '//Ŀ���� ù��° ITEM���� ����
'    oForm.ActiveItem = "CLTCOD"
    Call oForm.Items("DocDate").CLICK

    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY560_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY560_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY560_EnableMenus()
    
    On Error GoTo PH_PY560_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//����
    Call oForm.EnableMenu("1284", False)    '//���
    Call oForm.EnableMenu("1293", True)     '//�����
    
    Exit Sub
    
PH_PY560_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY560_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY560_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY560_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY560_FormItemEnabled
        Call PH_PY560_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY560_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY560_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY560_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY560_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY560_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", False)    '//�����߰�
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then

        Call oForm.EnableMenu("1281", False)    '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�
    ElseIf (oForm.Mode = fm_OK_MODE) Then

        
        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY560_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY560_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    Call PH_PY560_Print_Report01
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
                   Select Case pval.ItemUID
                   '//�ٹ����°� �ٲ�� �ٹ��� �缳��
                           Case "ShiftDat"
                                Set oCombo = oForm.Items("GNMUJO").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//����
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("ShiftDat").Specific.VALUE & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo)
                   '//������� �ٲ�� �μ��� ��� �缳��
                           Case "CLTCOD"
                                '//�μ�
                                Set oCombo = oForm.Items("TeamCode").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//����
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                '//���� ��������� �ٽ� Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                   End Select
                  
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
'                            '//����� �O�Ƽ� ȭ�� ǥ�� �ϱ�
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

                Call PH_PY560_FormItemEnabled
                Call PH_PY560_AddMatrixRow
                
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
'                Set oDS_PH_PY560A = Nothing
'                Set oDS_PH_PY560B = Nothing
                
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
                Call PH_PY560_FormItemEnabled
                Call PH_PY560_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY560_FormItemEnabled
                Call PH_PY560_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY560_FormItemEnabled
                Call PH_PY560_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY560_FormItemEnabled
            Case "1293" '// �����
'                '// [MAT1 ��]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY560B.Size - 1)
'                        If oDS_PH_PY560B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY560B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY560B.Size
'                        Call oDS_PH_PY560B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
                'End If
                Call PH_PY560_AddMatrixRow
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

Sub PH_PY560_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY560_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 ��]
    'oMat1.FlushToDataSource
    'oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY560B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY560B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY560B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY560B.Offset = oRow
'            oDS_PH_PY560B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY560B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY560B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY560B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY560B.Offset = oRow - 1
'            oDS_PH_PY560B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY560B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY560B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY560B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY560B.Offset = oRow
'        oDS_PH_PY560B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY560B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY560B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY560B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY560_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY560_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY560_FormClear()
On Error GoTo PH_PY560_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY560'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY560_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY560_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY560_DataValidCheck() As Boolean
On Error GoTo PH_PY560_DataValidCheck_Error
    PH_PY560_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    '//----------------------------------------------------------------------------------
'    '//�ʼ� üũ
'    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY560A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "�����ȣ�� �ʼ��Դϴ�.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY560_DataValidCheck = False
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix ������ �� ����(DB �����)
'    If oDS_PH_PY560B.Size > 1 Then oDS_PH_PY560B.RemoveRecord (oDS_PH_PY560B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY560_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY560_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY560_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY560_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY560_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY560_Validate_Error
    PH_PY560_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY560A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY560_Validate = False
        GoTo PH_PY560_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY560_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY560_Validate_Error:
    PH_PY560_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY560_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY560_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY560_Print_Report01_Error
    
    Dim CLTCOD          As String
    Dim DocDate         As String
    Dim TeamCode        As String
    
    Dim SukS            As String
    Dim YaS             As String
    Dim DangJ           As String
    Dim Chk1            As String
    Dim Chk2            As String
    Dim JoongS          As String
    Dim Bus1            As String
    Dim Bus2            As String
    Dim Bus3            As String
    Dim BusAdd          As String
    Dim G5_YN As String
    Dim NoAccdnt As String
    
    Dim oCheck          As SAPbouiCOM.CheckBox
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)
    
     '/ ODBC ���� üũ
    If ConnectODBC = False Then
        GoTo PH_PY560_Print_Report01_Error
    End If
    
    
    '//���� MOVE , Trim ��Ű��..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
    SukS = Trim(oForm.Items("SukS").Specific.VALUE)
    YaS = Trim(oForm.Items("YaS").Specific.VALUE)
    DangJ = Trim(oForm.Items("DangJik").Specific.VALUE)
    JoongS = Trim(oForm.Items("JoongS").Specific.VALUE)
    Bus1 = Trim(oForm.Items("Bus1").Specific.VALUE)
    Bus2 = Trim(oForm.Items("Bus2").Specific.VALUE)
    Bus3 = Trim(oForm.Items("Bus3").Specific.VALUE)
    BusAdd = Trim(oForm.Items("BusAdd").Specific.VALUE)
    
    If oForm.DataSources.UserDataSources.Item("Chk1").VALUE = "Y" Then
       Chk1 = "05:30���� "
    Else
        Chk1 = ""
    End If
    
    If oForm.DataSources.UserDataSources.Item("Chk2").VALUE = "Y" Then
       Chk2 = "08:30����"
    Else: Chk2 = ""
    End If
    
    If oForm.DataSources.UserDataSources.Item("G5_YN").VALUE = "Y" Then
        G5_YN = "Y"
    Else
        G5_YN = "N"
    End If
    
    If oForm.DataSources.UserDataSources.Item("NoAccdnt").VALUE = "Y" Then
        NoAccdnt = "Y"
    Else
        NoAccdnt = "N"
    End If
    
    
    '//����,����Check
    sQry = "Select b.U_DayType From [@PH_PY003A] a INNER JOIN [@PH_PY003B] b ON a.Code = b.Code WHERE a.U_CLTCOD = '" & CLTCOD & "' AND B.U_DATE = '" & DocDate & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields("U_DayType").VALUE <> "2" Then
          '------------------------------����-------------------------------------
            WinTitle = "[PH_PY560] �������Ȳ(����)"
            ReportName = "PH_PY560_01.rpt"
            '/ Formula �����ʵ�
            ReDim gRpt_Formula(2)
            ReDim gRpt_Formula_Value(2)
        
            gRpt_Formula(1) = "CLTCOD"
            sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
            Call oRecordSet.DoQuery(sQry)
            gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
        
            gRpt_Formula(2) = "DocDate"
            gRpt_Formula_Value(2) = Format(DocDate, "####-##-##")
        
        
            '/ SubReport
            ReDim gRpt_SRptSqry(3)
            ReDim gRpt_SRptName(3)
        
            ReDim gRpt_SFormula(3, 10)
            ReDim gRpt_SFormula_Value(3, 10)
        
            gRpt_SFormula(1, 1) = ""
            gRpt_SFormula_Value(1, 1) = ""
        
            gRpt_SFormula(2, 1) = "CLTCOD"
            gRpt_SFormula_Value(2, 1) = CLTCOD
        
            gRpt_SFormula(3, 1) = "SukS"
            gRpt_SFormula_Value(3, 1) = SukS
        
            gRpt_SFormula(3, 2) = "YaS"
            gRpt_SFormula_Value(3, 2) = YaS
        
            gRpt_SFormula(3, 3) = "DangJ"
            gRpt_SFormula_Value(3, 3) = DangJ
        
            gRpt_SFormula(3, 4) = "Chk1"
            gRpt_SFormula_Value(3, 4) = Chk1
        
            gRpt_SFormula(3, 5) = "Chk2"
            gRpt_SFormula_Value(3, 5) = Chk2
        
        
            sQry = "EXEC [PH_PY560_02] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "'"
            gRpt_SRptSqry(1) = sQry
            gRpt_SRptName(1) = "PH_PY560_SUB1"
        
            sQry = "EXEC [PH_PY560_03] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "'"
            gRpt_SRptSqry(2) = sQry
            gRpt_SRptName(2) = "PH_PY560_SUB2"
        
            sQry = "EXEC [PH_PY560_04] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "','" & G5_YN & "'"
            gRpt_SRptSqry(3) = sQry
            gRpt_SRptName(3) = "PH_PY560_SUB3"
        
        
            '/ Procedure ����"
            sQry = "EXEC [PH_PY560_01] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "'"
            
           '--------���ϳ�------------------------------------------------------------------------
   Else
            '------------------------------����-------------------------------------
             WinTitle = "[PH_PY560] �������Ȳ(����)"
             ReportName = "PH_PY560_05.rpt"
            
             '/ Formula �����ʵ�
             ReDim gRpt_Formula(10)
             ReDim gRpt_Formula_Value(10)
             
             gRpt_Formula(1) = "CLTCOD"
             sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
             Call oRecordSet.DoQuery(sQry)
             gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
             
             gRpt_Formula(2) = "DocDate"
             gRpt_Formula_Value(2) = Format(DocDate, "####-##-##")
             
             gRpt_Formula(3) = "JoongS"
             gRpt_Formula_Value(3) = JoongS
             
             gRpt_Formula(4) = "SukS"
             gRpt_Formula_Value(4) = SukS
             
             gRpt_Formula(5) = "YaS"
             gRpt_Formula_Value(5) = YaS
             
             gRpt_Formula(6) = "DangJ"
             gRpt_Formula_Value(6) = DangJ
             
             gRpt_Formula(7) = "Bus1"
             gRpt_Formula_Value(7) = Bus1
             
             gRpt_Formula(8) = "Bus2"
             gRpt_Formula_Value(8) = Bus2
             
             gRpt_Formula(9) = "Bus3"
             gRpt_Formula_Value(9) = Bus3
             
             gRpt_Formula(10) = "BusAdd"
             gRpt_Formula_Value(10) = BusAdd
             
             
             '/ SubReport
             ReDim gRpt_SRptSqry(1)
             ReDim gRpt_SRptName(1)
             
             ReDim gRpt_SFormula(1, 1)
             ReDim gRpt_SFormula_Value(1, 1)
             
             '/ Procedure ����"
             If NoAccdnt = "N" Then
                sQry = "EXEC [PH_PY560_05] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "'"
             Else
                sQry = "EXEC [PH_PY560_06] '" & CLTCOD & "', '" & DocDate & "', '" & TeamCode & "'"
             End If
             
            '--------���ϳ�------------------------------------------------------------------------
   End If
   
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY560_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V", , 1) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : ����!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY560_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY560_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub










