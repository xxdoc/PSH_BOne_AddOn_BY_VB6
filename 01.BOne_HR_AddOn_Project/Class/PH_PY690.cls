VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY690"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY690.cls
'//  Module         : PH
'//  Desc           : ��������Ȳ
'//  �ۼ���         : NGY
'//  DATE           : 2012.12.27
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// �׸��� ����
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PY690     As SAPbouiCOM.DataTable
'
''// ��Ʈ���� ����
'Public oMat1 As SAPbouiCOM.Matrix
'Private oDS_PH_PY690A As SAPbouiCOM.DBDataSource
'Private oDS_PH_PY690B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY690.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY690_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY690"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY690_CreateItems
    Call PH_PY690_EnableMenus
    Call PH_PY690_SetDocument(oFromDocEntry01)
'    Call PH_PY690_FormResize
    
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

Private Function PH_PY690_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY690_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// ������ ����
    '//----------------------------------------------------------------------------------------------
    
    '//�����
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// �����ڿ� ���� ����� ����
    oForm.DataSources.UserDataSources.Item("CLTCOD").Value = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
    
    '//���ؿ�
    oForm.Items("MM").Specific.String = Format(Now, "MM")
    
    '//Ŀ���� ù��° ITEM���� ����
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
    
PH_PY690_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY690_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY690_EnableMenus()
    
    On Error GoTo PH_PY690_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//����
    Call oForm.EnableMenu("1284", False)    '//���
    Call oForm.EnableMenu("1293", True)     '//�����
    
    Exit Sub
    
PH_PY690_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY690_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY690_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY690_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY690_FormItemEnabled
        Call PH_PY690_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY690_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY690_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY690_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY690_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY690_FormItemEnabled_Error
    
    
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
    
PH_PY690_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY690_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUid = "Btn01" Then
                    Call PH_PY690_Print_Report01
                End If
            ElseIf pval.BeforeAction = False Then


            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUid
            Case "Mat01", "Grid01"
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
                    Select Case pval.ItemUid
                           
                           
                    End Select
                End If
            End If
            
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUid
                Case "Mat01"
'                    If pval.Row > 0 Then
'                        Call oMat1.SelectRow(pval.Row, True, False)
'                    End If
                End Select
                
                Select Case pval.ItemUid
                Case "Mat01", "Grid01"
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
       
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUid
'                        Case "MSTCOD"
'                            '//����� �O�Ƽ� ȭ�� ǥ�� �ϱ�
'                            sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" & Trim(oForm.Items("MSTCOD").Specific.Value) & "'"
'                            oRecordSet.DoQuery sQry
'                            oForm.Items("MSTNAME").Specific.String = Trim(oRecordSet.Fields("U_FullName").Value)
                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
               'oMat1.LoadFromDataSource

                Call PH_PY690_FormItemEnabled
                Call PH_PY690_AddMatrixRow
                
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
'                Set oDS_PH_PY690A = Nothing
'                Set oDS_PH_PY690B = Nothing
                
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
                Call PH_PY690_FormItemEnabled
                Call PH_PY690_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY690_FormItemEnabled
                Call PH_PY690_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY690_FormItemEnabled
                Call PH_PY690_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY690_FormItemEnabled
            Case "1293" '// �����
'                '// [MAT1 ��]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY690B.Size - 1)
'                        If oDS_PH_PY690B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY690B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY690B.Size
'                        Call oDS_PH_PY690B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
                'End If
                Call PH_PY690_AddMatrixRow
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

Sub PH_PY690_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY690_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 ��]
    'oMat1.FlushToDataSource
    'oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY690B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY690B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY690B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY690B.Offset = oRow
'            oDS_PH_PY690B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY690B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY690B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY690B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY690B.Offset = oRow - 1
'            oDS_PH_PY690B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY690B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY690B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY690B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY690B.Offset = oRow
'        oDS_PH_PY690B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY690B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY690B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY690B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY690_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY690_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY690_FormClear()
On Error GoTo PH_PY690_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY690'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY690_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY690_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY690_DataValidCheck() As Boolean
On Error GoTo PH_PY690_DataValidCheck_Error
    PH_PY690_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    '//----------------------------------------------------------------------------------
'    '//�ʼ� üũ
'    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY690A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "�����ȣ�� �ʼ��Դϴ�.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY690_DataValidCheck = False
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix ������ �� ����(DB �����)
'    If oDS_PH_PY690B.Size > 1 Then oDS_PH_PY690B.RemoveRecord (oDS_PH_PY690B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY690_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY690_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY690_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY690_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY690_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY690_Validate_Error
    PH_PY690_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY690A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY690_Validate = False
        GoTo PH_PY690_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY690_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY690_Validate_Error:
    PH_PY690_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY690_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY690_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY690_Print_Report01_Error
    
    Dim CLTCOD          As String
    Dim MM              As String
    
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
     '/ ODBC ���� üũ
    If ConnectODBC = False Then
        GoTo PH_PY690_Print_Report01_Error
    End If
    
    
    '//���� MOVE , Trim ��Ű��..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.Value)
    MM = Trim(oForm.Items("MM").Specific.Value)
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY690] ��������Ȳ"
    ReportName = "PH_PY690_01.rpt"
    
    '/ Formula �����ʵ�
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
    gRpt_Formula(1) = "CLTCOD"
    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).Value
    
    gRpt_Formula(2) = "MM"
    gRpt_Formula_Value(2) = MM
    
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure ����"
    sQry = "EXEC [PH_PY690_01] '" & CLTCOD & "', '" & MM & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PH_PY690_Print_Report01_Error
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 1) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : ����!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY690_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY690_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub


















