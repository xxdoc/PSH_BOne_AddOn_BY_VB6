VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY117.cls
'//  Module         : �޿����� > �޻󿩸����۾�
'//  Desc           : �޻󿩸����۾�
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

'// �׸��� ����
Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY117     As SAPbouiCOM.DataTable


Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Private sCLTCOD     As String
Private sYM         As String
Private sJOBTYP     As String
Private sJOBGBN     As String
Private sPAYSEL     As String
Private sMSTCOD     As String
Private sTeamCode   As String
Private sRspCode    As String
Private sClose      As String

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY117.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY117_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY117"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY117_CreateItems
    Call PH_PY117_EnableMenus
    Call PH_PY117_SetDocument(oFromDocEntry01)
'    Call PH_PY117_FormResize
    
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

Private Function PH_PY117_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY117_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// �����ͼ�����
    '//----------------------------------------------------------------------------------------------

    '//���̺��� ���°�� �����ͼ�(Grid)
    oForm.DataSources.DataTables.Add ("PH_PY117")
    oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add "����", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add "�μ�", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add "���", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add "���", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add "����", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add "�����޾�", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY117").Columns.Add "�����޾�", ft_AlphaNumeric

    Set oGrid1 = oForm.Items("Grid1").Specific

    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY117")
    Set oDS_PH_PY117 = oForm.DataSources.DataTables.Item("PH_PY117")
    
    
    '//----------------------------------------------------------------------------------------------
    '// ������ ����
    '//----------------------------------------------------------------------------------------------

'    '//UserDataSources
'    Call oForm.DataSources.UserDataSources.Add("     ", dt_SHORT_TEXT, 10)
'    Set oCombo = oForm.Items("    ").Specific
'    oCombo.DataBind.SetBound True, "", "    "
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//�μ�
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("TeamCode").Specific
    oCombo.DataBind.SetBound True, "", "TeamCode"
    
    '//���
    Call oForm.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("RspCode").Specific
    oCombo.DataBind.SetBound True, "", "RspCode"
    
    '//�����
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 1)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    Call CLTCOD_Select(oForm, "CLTCOD")
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// �ͼӳ��
    Call oForm.DataSources.UserDataSources.Add("YM", dt_SHORT_TEXT, 6)
    Set oEdit = oForm.Items("YM").Specific
    oEdit.DataBind.SetBound True, "", "YM"
    oForm.DataSources.UserDataSources.Item("YM").ValueEx = Format(Date, "YYYYMM")
'    Call oDS_PH_PY117.setValue("U_YM", 0, Format(Date, "YYYYMM"))
    
    
    '//��������
    Call oForm.DataSources.UserDataSources.Add("JOBTYP", dt_SHORT_TEXT, 1)
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.DataBind.SetBound True, "", "JOBTYP"
    oCombo.ValidValues.Add "1", "�޿�"
    oCombo.ValidValues.Add "2", "��"
    oCombo.Select "0", psk_Index
    oForm.Items("JOBTYP").DisplayDesc = True
    
    '//���ޱ���
    Call oForm.DataSources.UserDataSources.Add("JOBGBN", dt_SHORT_TEXT, 1)
    Set oCombo = oForm.Items("JOBGBN").Specific
    oCombo.DataBind.SetBound True, "", "JOBGBN"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.Select "0", psk_Index
    oForm.Items("JOBGBN").DisplayDesc = True
    
    '//���޴��
    Call oForm.DataSources.UserDataSources.Add("PAYSEL", dt_SHORT_TEXT, 1)
    Set oCombo = oForm.Items("PAYSEL").Specific
    oCombo.DataBind.SetBound True, "", "PAYSEL"
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "��ü"
    oCombo.Select "%", psk_ByValue
    oForm.Items("PAYSEL").DisplayDesc = True
    
    '//���
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    
    '// ����
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 20)
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    
    '//����
    Call oForm.DataSources.UserDataSources.Add("Close", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("Close").Specific
    oCheck.DataBind.SetBound True, "", "Close"
    oCheck.ValOn = "Y"
    oCheck.ValOff = "N"
    oCheck.Checked = False
    
    
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY117_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY117_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY117_EnableMenus()
    
    On Error GoTo PH_PY117_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//����
    Call oForm.EnableMenu("1284", False)    '//���
    Call oForm.EnableMenu("1293", True)     '//�����
    
    Exit Sub
    
PH_PY117_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY117_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY117_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY117_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY117_FormItemEnabled
        Call PH_PY117_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY117_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY117_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY117_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY117_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY117_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        Call oForm.EnableMenu("1281", False)     '//����ã��
        Call oForm.EnableMenu("1282", False)    '//�����߰�
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then

        Call oForm.EnableMenu("1281", False)    '//����ã��
        Call oForm.EnableMenu("1282", False)     '//�����߰�
    ElseIf (oForm.Mode = fm_OK_MODE) Then

        
        Call oForm.EnableMenu("1281", False)     '//����ã��
        Call oForm.EnableMenu("1282", False)     '//�����߰�

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY117_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY117_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUid = "Btn1" Then
                    If Sbo_Application.MessageBox("����ó���� ���� �Ͻðڽ��ϱ�?", 2, "Yes", "No") = 2 Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call PH_PY117_DataChange
                    Call PH_PY117_DataSearch
                End If
                If pval.ItemUid = "Btn_Search" Then
                    Call PH_PY117_DataSearch
                End If
                If pval.ItemUid = "Btn_All" Then
                    oForm.Freeze True
                    For i = 0 To oGrid1.Rows.Count - 1
                        oDS_PH_PY117.setValue "MAGAM", i, "Y"
                    Next
                    oForm.Freeze False
                End If
                If pval.ItemUid = "Btn_Rev" Then
                    oForm.Freeze True
                    For i = 0 To oGrid1.Rows.Count - 1
                        oDS_PH_PY117.setValue "MAGAM", i, "N"
                    Next
                    oForm.Freeze False
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUid
            Case "Mat1", "Grid1"
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
                    If pval.ItemUid = "CLTCOD" Then
                        
                        '//�⺻���� - �μ� (����忡 ���� �μ�����)
                        Set oCombo = oForm.Items("TeamCode").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
                        sQry = sQry & " ORDER BY U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        oCombo.ValidValues.Add "%", "��ü"
                        oCombo.Select "%", psk_ByValue
                        oForm.Items("TeamCode").DisplayDesc = True
                        
                        '//��� (����忡 ���� ��纯��)
                    
                        Set oCombo = oForm.Items("RspCode").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
                        sQry = sQry & " Order By U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        oCombo.ValidValues.Add "%", "��ü"
                        oCombo.Select "%", psk_ByValue
                        oForm.Items("RspCode").DisplayDesc = True
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6

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
                    If pval.ItemUid = "MSTCOD" Then
                        If oForm.Items("MSTCOD").Specific.Value = "" Then
                            oForm.Items("MSTNAM").Specific.Value = ""
                        Else
                            oForm.Items("MSTNAM").Specific.Value = MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD").Specific.Value & "'")
                        End If
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
'            If pval.BeforeAction = True Then
'            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'
'                Call PH_PY117_FormItemEnabled
'                Call PH_PY117_AddMatrixRow
'
'            End If
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
                Set oDS_PH_PY117 = Nothing

                Set oGrid1 = Nothing
                
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
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.Before_Action = False Then
'
'            End If
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
                Call PH_PY117_FormItemEnabled
                Call PH_PY117_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY117_FormItemEnabled
                Call PH_PY117_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY117_FormItemEnabled
                Call PH_PY117_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY117_FormItemEnabled
            Case "1293" '// �����
'                '// [MAT1 ��]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY117B.Size - 1)
'                        If oDS_PH_PY117B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY117B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY117B.Size
'                        Call oDS_PH_PY117B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
'                Call PH_PY117_AddMatrixRow
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
    Case "Mat1"
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

Sub PH_PY117_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY117_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 ��]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY117B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY117B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY117B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY117B.Offset = oRow
'            oDS_PH_PY117B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY117B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY117B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY117B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY117B.Offset = oRow - 1
'            oDS_PH_PY117B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY117B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY117B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY117B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY117B.Offset = oRow
'        oDS_PH_PY117B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY117B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY117B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY117B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY117_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY117_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY117_FormClear()
On Error GoTo PH_PY117_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY117'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY117_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY117_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY117_DataValidCheck() As Boolean
On Error GoTo PH_PY117_DataValidCheck_Error
    PH_PY117_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//�ʼ� üũ
    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY117A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "�����ȣ�� �ʼ��Դϴ�.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix ������ �� ����(DB �����)
'    If oDS_PH_PY117B.Size > 1 Then oDS_PH_PY117B.RemoveRecord (oDS_PH_PY117B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY117_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY117_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY117_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY117_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY117_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY117_Validate_Error
    PH_PY117_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY117A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY117_Validate = False
        GoTo PH_PY117_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY117_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY117_Validate_Error:
    PH_PY117_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY117_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY117_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY117_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC ���� üũ
    If ConnectODBC = False Then
        GoTo PH_PY117_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] ���ּ�"
    ReportName = "S142_1.rpt"
    sQry = "EXEC PH_PY117_1 '" & oForm.Items("8").Specific.Value & "'"
    
    '/ Formula �����ʵ�
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure ����"
    sQry = "EXEC [PS_PP820_01] '" & BPLID & "', '" & ItmBsort & "', '" & DocDate & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
            Sbo_Application.SetStatusBarMessage "gCryReport_Action : ����!", bmt_Short, True
        End If
    Else
        Sbo_Application.SetStatusBarMessage "��ȸ�� �����Ͱ� �����ϴ�.", bmt_Short, True
    End If
    
    Exit Sub
    
PH_PY117_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY117_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY117_DataSearch()
    Dim i           As Long
    Dim iRow        As Long
    Dim sQry        As String
    
    
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY117_DataSearch_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    oForm.Freeze (True)
    
    sCLTCOD = Trim(oForm.Items("CLTCOD").Specific.Value)
    sYM = Trim(oForm.Items("YM").Specific.Value)
    sJOBTYP = Trim(oForm.Items("JOBTYP").Specific.Value)
    sJOBGBN = Trim(oForm.Items("JOBGBN").Specific.Value)
    sPAYSEL = Trim(oForm.Items("PAYSEL").Specific.Value)
    sMSTCOD = Trim(oForm.Items("MSTCOD").Specific.Value)
    If sMSTCOD = "" Then sMSTCOD = "%"
    sTeamCode = Trim(oForm.Items("TeamCode").Specific.Value)
    sRspCode = Trim(oForm.Items("RspCode").Specific.Value)
    sClose = Trim(oForm.DataSources.UserDataSources.Item("Close").ValueEx)
    If sClose = "" Then sClose = "N"
    
    sQry = "Exec PH_PY117 '" & sCLTCOD & "','" & sYM & "','" & sJOBTYP & "','" & sJOBGBN & "','" & sPAYSEL & "','" & sMSTCOD & "',"
    sQry = sQry & "'" & sTeamCode & "','" & sRspCode & "','" & sClose & "'"
    
    Call oDS_PH_PY117.ExecuteQuery(sQry)
    
    Debug.Print oDS_PH_PY117.Rows.Count
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY117_TitleSetting(iRow)
    
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Exit Sub
    
PH_PY117_DataSearch_Error:
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY004_DataFind_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY117_DataChange()
    Dim i           As Long
    Dim iRow        As Long
    Dim sQry        As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY117_DataChange_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    oForm.Freeze (True)
    
    For i = 0 To oGrid1.Rows.Count - 1
        If oDS_PH_PY117.GetValue("MAGAM", i) = "Y" Then
            sQry = "UPDATE [@PH_PY112A] SET U_ENDCHK = 'Y' WHERE U_MSTCOD = '" & oDS_PH_PY117.GetValue("U_MSTCOD", i) & "'"
            sQry = sQry & " AND U_CLTCOD = '" & sCLTCOD & "'"
            sQry = sQry & " AND U_YM = '" & sYM & "'"
            sQry = sQry & " AND U_JOBTYP = '" & sJOBTYP & "'"
            sQry = sQry & " AND U_JOBGBN = '" & sJOBGBN & "'"
            sQry = sQry & " AND (U_JOBTRG = '" & sPAYSEL & "' OR (U_JOBTRG <> '" & sPAYSEL & "' AND U_JOBTRG LIKE '" & sPAYSEL & "'))"
            
            oRecordSet.DoQuery sQry
        End If
    Next
    
    Sbo_Application.StatusBar.SetText "����ó���� ���� �Ǿ����ϴ�.", bmt_Short, smt_Success
    
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Exit Sub
    
PH_PY117_DataChange_Error:
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY004_DataFind_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY117_TitleSetting(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(80)       As String

    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "����"
    COLNAM(1) = "�μ�"
    COLNAM(2) = "���"
    COLNAM(3) = "���"
    COLNAM(4) = "����"
    COLNAM(5) = "�����޾�"
    COLNAM(6) = "�Ѱ�����"
    COLNAM(7) = "�����޾�"
    For i = 1 To 36
        COLNAM(i + 7) = "�����׸�" & i
    Next i
    For i = 1 To 36
        COLNAM(i + 43) = "�����׸�" & i
    Next i

    For i = 0 To UBound(COLNAM) - 1
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        oGrid1.Columns.Item(i).Editable = False
        If i = 0 Then
            If sClose = "N" Then
                oGrid1.Columns.Item(i).Editable = True
            End If
            oGrid1.Columns.Item(i).Type = gct_CheckBox
        End If
    Next i
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY117_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub