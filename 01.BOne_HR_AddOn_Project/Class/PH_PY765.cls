VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY765"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY765.cls
'//  Module         : PH
'//  Desc           : �޿�����������
'//  �ۼ���         : NGY
'//  DATE           : 2013.03.20
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// �׸��� ����
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PY765     As SAPbouiCOM.DataTable
'
''// ��Ʈ���� ����
'Public oMat1 As SAPbouiCOM.Matrix
'Private oDS_PH_PY765A As SAPbouiCOM.DBDataSource
'Private oDS_PH_PY765B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY765.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY765_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY765"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY765_CreateItems
    Call PH_PY765_EnableMenus
    Call PH_PY765_SetDocument(oFromDocEntry01)
'    Call PH_PY765_FormResize
    
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

Private Function PH_PY765_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY765_CreateItems_Error
    
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
    
    '//���޳��
    oForm.Items("YM").Specific.String = Format(Now, "YYYYMM")
    
    '//Ŀ���� ù��° ITEM���� ����

    '���������
    Call oForm.DataSources.UserDataSources.Add("Chk01", dt_SHORT_TEXT, 1)
    Call oForm.Items("Chk01").Specific.DataBind.SetBound(True, "", "Chk01")
    oForm.Items("Chk01").Specific.Checked = False
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    
'    oForm.ActiveItem = "CLTCOD"
    
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY765_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY765_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY765_EnableMenus()
    
    On Error GoTo PH_PY765_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//����
    Call oForm.EnableMenu("1284", False)    '//���
    Call oForm.EnableMenu("1293", True)     '//�����
    
    Exit Sub
    
PH_PY765_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY765_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY765_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY765_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY765_FormItemEnabled
        Call PH_PY765_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY765_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY765_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY765_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY765_FormItemEnabled()
On Error GoTo PH_PY765_FormItemEnabled_Error
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    
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
    
PH_PY765_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY765_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error

    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "Btn01" Then
                    If PH_PY765_DataValidCheck() = True Then
                        Call PH_PY765_Print_Report01
                    End If
                    
                End If
            ElseIf pval.BeforeAction = False Then


            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.CharPressed = 9 Then
                If pval.ItemUID = "JIGBIL" Then
                    If Trim(oForm.Items("JIGBIL").Specific.VALUE) = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "MSTCOD" Then
                    If oForm.Items("MSTCOD").Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            
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

                Call PH_PY765_FormItemEnabled
                Call PH_PY765_AddMatrixRow
                
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
'                Set oDS_PH_PY765A = Nothing
'                Set oDS_PH_PY765B = Nothing
                
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
On Error GoTo Raise_FormMenuEvent_Error

    Dim i As Long
    
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
                Call PH_PY765_FormItemEnabled
                Call PH_PY765_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY765_FormItemEnabled
                Call PH_PY765_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY765_FormItemEnabled
                Call PH_PY765_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY765_FormItemEnabled
            Case "1293" '// �����
'                '// [MAT1 ��]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY765B.Size - 1)
'                        If oDS_PH_PY765B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY765B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY765B.Size
'                        Call oDS_PH_PY765B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
                'End If
                Call PH_PY765_AddMatrixRow
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

Sub PH_PY765_AddMatrixRow()
On Error GoTo PH_PY765_AddMatrixRow_Error

    Dim oRow    As Long
    
    Call oForm.Freeze(True)
    
    '//[Mat1 ��]
    'oMat1.FlushToDataSource
    'oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY765B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY765B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY765B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY765B.Offset = oRow
'            oDS_PH_PY765B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY765B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY765B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY765B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY765B.Offset = oRow - 1
'            oDS_PH_PY765B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY765B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY765B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY765B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY765B.Offset = oRow
'        oDS_PH_PY765B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY765B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY765B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY765B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY765_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY765_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY765_FormClear()
On Error GoTo PH_PY765_FormClear_Error

    Dim DocEntry As String
    
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY765'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY765_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY765_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY765_DataValidCheck() As Boolean
On Error GoTo PH_PY765_DataValidCheck_Error

    PH_PY765_DataValidCheck = False
    
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//�ʼ� üũ
    '//----------------------------------------------------------------------------------
    If Trim(oForm.Items("YM").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "���޳���� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        PH_PY765_DataValidCheck = False
        Exit Function
    End If
    
    PH_PY765_DataValidCheck = True
    
    Set oRecordSet = Nothing
    
    Exit Function
    
PH_PY765_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY765_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY765_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY765_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY765_Validate_Error

    PH_PY765_Validate = True
    
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY765A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY765_Validate = False
        GoTo PH_PY765_Validate_Exit
    End If

    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY765_Validate_Exit:

    Set oRecordSet = Nothing
    
PH_PY765_Validate_Error:
    PH_PY765_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY765_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY765_Print_Report01()
On Error GoTo PH_PY765_Print_Report01_Error

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim CLTCOD          As String
    Dim YM              As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)
    
     '/ ODBC ���� üũ
    If ConnectODBC = False Then
        GoTo PH_PY765_Print_Report01_Error
    End If
    
    '//���� MOVE , Trim ��Ű��..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    YM = Trim(oForm.Items("YM").Specific.VALUE)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY765] �޿�����������"
    
    '/ Formula �����ʵ�
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
    gRpt_Formula(1) = "CLTCOD"
    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
    
    gRpt_Formula(2) = "YM"
    gRpt_Formula_Value(2) = Format(YM, "####-##")
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    If oForm.DataSources.UserDataSources.Item("Chk01").VALUE = "Y" Then '��������� ����
    
        ReportName = "PH_PY765_02.rpt"
        sQry = "EXEC [PH_PY765_02] '" & CLTCOD & "', '" & YM & "'"

    Else
        
        ReportName = "PH_PY765_01.rpt"
        sQry = "EXEC [PH_PY765_01] '" & CLTCOD & "', '" & YM & "'"

    End If
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY765_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 1) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : ����!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY765_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY765_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub
