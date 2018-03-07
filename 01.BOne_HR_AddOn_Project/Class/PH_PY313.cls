VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY313"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY313.cls
'//  Module         : �λ���� > ��Ÿ > ��αݰ���
'//  Desc           : ��αݰ��
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY313A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY313B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY313.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY313_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY313"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY313_CreateItems
    Call PH_PY313_EnableMenus
    Call PH_PY313_SetDocument(oFromDocEntry01)
'    Call PH_PY313_FormResize
    
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

Private Function PH_PY313_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY313_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY313A = oForm.DataSources.DBDataSources("@PH_PY313A")
    Set oDS_PH_PY313B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    Set oMat1 = oForm.Items("Mat01").Specific

    oMat1.SelectionMode = ms_Auto
    oMat1.AutoResizeColumns

    '�����_S
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 5)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    '�����_E
    
    '��ȯ����_S
    Call oForm.DataSources.UserDataSources.Add("RpmtDate", dt_DATE)
    Call oForm.Items("RpmtDate").Specific.DataBind.SetBound(True, "", "RpmtDate")
    '��ȯ����_E
    
    '���_S
    Call oForm.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("CntcCode").Specific.DataBind.SetBound(True, "", "CntcCode")
    '���_E
    
    '����_S
    Call oForm.DataSources.UserDataSources.Add("CntcName", dt_SHORT_TEXT, 50)
    Call oForm.Items("CntcName").Specific.DataBind.SetBound(True, "", "CntcName")
    '����_E

    '��Ͽ���_S
    Call oForm.DataSources.UserDataSources.Add("RegYN", dt_SHORT_TEXT, 50)
    Call oForm.Items("RegYN").Specific.DataBind.SetBound(True, "", "RegYN")
    '��Ͽ���_E

    '//----------------------------------------------------------------------------------------------
    '// �⺻����
    '//----------------------------------------------------------------------------------------------
    
    '�����
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '��Ͽ���
    Set oCombo = oForm.Items("RegYN").Specific
    Call oCombo.ValidValues.Add("Y", "Y")
    Call oCombo.ValidValues.Add("N", "N")
    Call oCombo.Select("N", psk_ByValue) '"N" �⺻ ����
    oForm.Items("RegYN").DisplayDesc = True
    
    '��Ʈ����-��Ͽ���
    Set oColumn = oMat1.Columns("RegYN")
    Call oColumn.ValidValues.Add("Y", "Y")
    Call oColumn.ValidValues.Add("N", "N")
    oColumn.DisplayDesc = True
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY313_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY313_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY313_EnableMenus()
    
    On Error GoTo PH_PY313_EnableMenus_Error
    
'    Call oForm.EnableMenu("1283", False)         '// ����
'    Call oForm.EnableMenu("1287", True)          '// ����
''    Call oForm.EnableMenu("1286", True)         '// �ݱ�
'    Call oForm.EnableMenu("1284", True)         '// ���
'    Call oForm.EnableMenu("1293", True)         '// �����
    
    Exit Sub
    
PH_PY313_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY313_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY313_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY313_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY313_FormItemEnabled
'        Call PH_PY313_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY313_FormItemEnabled
        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY313_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY313_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY313_FormItemEnabled()
    On Error GoTo PH_PY313_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim CLTCOD As String
        
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        '�� DocEntry ����
'        Call PH_PY313_FormClear
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        oForm.Items("RpmtDate").Specific.Value = Format(Date, "YYYYMMDD")
        
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
    
PH_PY313_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY313_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry As String
    Dim i As Long
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Dim loopCount As Integer 'For Loop ��
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If PH_PY313_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'
'                        '//�ؾ����� �۾�
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If PH_PY313_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'                        '//�ؾ����� �۾�
'
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                    End If
'                End If
                If pval.ItemUID = "btnSearch" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY313_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'
'                        '//�ؾ����� �۾�
                        Call PH_PY313_MTX01
                        
                    End If
                
                ElseIf pval.ItemUID = "btnReg" Then
                
                    Dim FirstYN As String
                
                    Call oForm.Freeze(True)
                
                    oMat1.FlushToDataSource
                    FirstYN = Trim(oDS_PH_PY313B.GetValue("U_ColReg07", 0))
                    For loopCount = 1 To oMat1.VisualRowCount - 1
                        
                        Call oDS_PH_PY313B.setValue("U_ColReg07", loopCount, FirstYN)
                        
                    Next
                    oMat1.LoadFromDataSource
                    
                    Call oForm.Freeze(False)
                    
                ElseIf pval.ItemUID = "btnAdd" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                    
                        If PH_PY313_AddData = False Then '��α� ��� ������ �Է�
                            BubbleEvent = False
                            Exit Sub
                        Else
                            Call PH_PY313_MTX01 '��� �� �� ��ȸ
                            Call Sbo_Application.StatusBar.SetText("��α� ��� ������ ��ϵǾ����ϴ�.", bmt_Short, smt_Success)
                        End If
                        
                    End If
                
                End If
                
            ElseIf pval.BeforeAction = False Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY313_FormItemEnabled
'                            Call PH_PY313_AddMatrixRow
'                        End If
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY313_FormItemEnabled
'                            Call PH_PY313_AddMatrixRow
'                        End If
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY313_FormItemEnabled
'                        End If
'                    End If
'                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                If pval.ItemUID = "Mat01" Then
        
                    If pval.ColUID = "Name" And pval.CharPressed = "9" Then
                    
                        If oMat1.Columns.Item("Name").Cells(pval.Row).Specific.Value = "" Then
                            Call Sbo_Application.ActivateMenuItem("7425")
                            BubbleEvent = False
                        End If
                    
                    End If
                
                ElseIf pval.ItemUID = "CntcCode" And pval.CharPressed = "9" Then
                
                    If oForm.Items("CntcCode").Specific.Value = "" Then
                        Call Sbo_Application.ActivateMenuItem("7425")
                        BubbleEvent = False
                    End If
            
                End If

            ElseIf pval.Before_Action = False Then

            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
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
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
'                    Call PH_PY313_AddMatrixRow
                    Call oMat1.AutoResizeColumns
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
                End Select
                
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
                    
                    Select Case pval.ItemUID
                    
                        Case "CntcCode"
                        
'                            oDS_PH_PY309A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUid).Specific.Value & "'"))
                            oForm.Items("CntcName").Specific.Value = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.Value & "'")
                        
                        Case "Mat01"
                        
                            If pval.ColUID = "RpmtAmt" Then '��ȯ�ݾ� �Է� ��

                                Dim PrRmainAmt As Double '������ȯ�ܾ�
                                Dim RpmtAmt As Double '��ȯ�ݾ�
                                Dim RmainAmt As Double '��ȯ�ܾ�
                                
                                oMat1.FlushToDataSource

                                PrRmainAmt = oDS_PH_PY313B.GetValue("U_ColSum02", pval.Row - 1)
                                RpmtAmt = oDS_PH_PY313B.GetValue("U_ColSum03", pval.Row - 1)
                            
                                RmainAmt = PrRmainAmt - RpmtAmt
                            
                                Call oDS_PH_PY313B.setValue("U_ColSum05", pval.Row - 1, RmainAmt)

                                oMat1.LoadFromDataSource

'                                If oMat1.RowCount = pval.Row And Trim(oDS_PH_PY309B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                                    Call PH_PY309_AddMatrixRow
'                                End If

                            End If

                            Call oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                            Call oMat1.AutoResizeColumns
                            
                    End Select

                End If
                
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY313_FormItemEnabled
                Call PH_PY313_AddMatrixRow
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
                Set oDS_PH_PY313A = Nothing
                Set oDS_PH_PY313B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY313A", "Code")
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
            
            Case "7169": '���� ��������
            
                '���� �������� ���� �� ��Ʈ������ ���� ������ �࿡ �� �� �߰�
                Call PH_PY313_AddMatrixRow
            
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY313_FormItemEnabled
                Call PH_PY313_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY313_FormItemEnabled
                Call PH_PY313_AddMatrixRow
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY313_FormItemEnabled
                Call PH_PY313_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY313_FormItemEnabled
            Case "1293" '// �����

                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY313B.Size - 1)
                        If oDS_PH_PY313B.GetValue("U_LineNum", i) = "" Then
                            oDS_PH_PY313B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY313B.Size
                        Call oDS_PH_PY313B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY313_AddMatrixRow
                
            Case "7169": '���� ��������
            
                '���� �������� ���� ó��
                Call oForm.Freeze(True)
                Call oDS_PH_PY313B.RemoveRecord(oDS_PH_PY313B.Size - 1)
                Call oMat1.LoadFromDataSource
                Call oForm.Freeze(False)
                
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

Sub PH_PY313_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY313_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY313B.GetValue("U_LineNum", oRow - 1)) <> "" Then
            If oDS_PH_PY313B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY313B.InsertRecord (oRow)
            End If
            oDS_PH_PY313B.Offset = oRow
            oDS_PH_PY313B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY313B.setValue "U_ColReg01", oRow, ""
            oDS_PH_PY313B.setValue "U_ColReg02", oRow, ""
            oDS_PH_PY313B.setValue "U_ColReg03", oRow, ""
            oDS_PH_PY313B.setValue "U_ColSum01", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY313B.Offset = oRow - 1
            oDS_PH_PY313B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY313B.setValue "U_ColReg01", oRow - 1, ""
            oDS_PH_PY313B.setValue "U_ColReg02", oRow - 1, ""
            oDS_PH_PY313B.setValue "U_ColReg03", oRow - 1, ""
            oDS_PH_PY313B.setValue "U_ColSum01", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY313B.Offset = oRow
        oDS_PH_PY313B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY313B.setValue "U_ColReg01", oRow, ""
        oDS_PH_PY313B.setValue "U_ColReg02", oRow, ""
        oDS_PH_PY313B.setValue "U_ColReg03", oRow, ""
        oDS_PH_PY313B.setValue "U_ColSum01", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY313_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY313_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY313_FormClear()
On Error GoTo PH_PY313_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY313'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY313_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY313_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY313_DataValidCheck() As Boolean
On Error GoTo PH_PY313_DataValidCheck_Error
    PH_PY313_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '�����
    If Trim(oForm.Items("CLTCOD").Specific.Value) = "" Then
        Sbo_Application.SetStatusBarMessage "������� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY313_DataValidCheck = False
        Exit Function
    End If
    
    '��ȯ����
    If Trim(oForm.Items("RpmtDate").Specific.Value) = "" Then
        Sbo_Application.SetStatusBarMessage "��ȯ���ڴ� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("RpmtDate").CLICK ct_Regular
        PH_PY313_DataValidCheck = False
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
'                PH_PY313_DataValidCheck = False
'                Exit Function
'            End If
'
'            '�б���
'            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "�б����� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
'                PH_PY313_DataValidCheck = False
'                Exit Function
'            End If
'
'            '�г�
'            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "�г��� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
'                PH_PY313_DataValidCheck = False
'                Exit Function
'            End If
'
'            'ȸ��
'            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "ȸ���� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
'                PH_PY313_DataValidCheck = False
'                Exit Function
'            End If
'
'        Next
'    Else
'        Sbo_Application.SetStatusBarMessage "���� �����Ͱ� �����ϴ�.", bmt_Short, True
'        PH_PY313_DataValidCheck = False
'        Exit Function
'    End If
'
'        oMat1.FlushToDataSource
'    '// Matrix ������ �� ����(DB �����)
'    If oDS_PH_PY313B.Size > 1 Then oDS_PH_PY313B.RemoveRecord (oDS_PH_PY313B.Size - 1)
'
'    oMat1.LoadFromDataSource
    
    PH_PY313_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY313_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY313_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY313_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY313_MTX01()
On Error GoTo PH_PY313_MTX01_Error

    '//��Ʈ������ ������ �ε�
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("CLTCOD").Specific.Value
    Param02 = oForm.Items("RpmtDate").Specific.Value
    Param03 = oForm.Items("CntcCode").Specific.Value
    Param04 = oForm.Items("RegYN").Specific.Value
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet.RecordCount, False)
    
    sQry = "EXEC PH_PY313_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call oMat1.Clear
        GoTo PH_PY313_MTX01_Exit
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY313B.InsertRecord (i)
        End If
        oDS_PH_PY313B.Offset = i
        Call oDS_PH_PY313B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY313B.setValue("U_ColReg01", i, oRecordSet.Fields("LoanDoc").Value) '��αݹ�����ȣ
        Call oDS_PH_PY313B.setValue("U_ColReg02", i, oRecordSet.Fields("CntcCode").Value) '���
        Call oDS_PH_PY313B.setValue("U_ColReg03", i, oRecordSet.Fields("CntcName").Value) '����
        Call oDS_PH_PY313B.setValue("U_ColReg04", i, oRecordSet.Fields("LoanDate").Value) '��������
        Call oDS_PH_PY313B.setValue("U_ColSum01", i, oRecordSet.Fields("LoanAmt").Value) '����ݾ�
        Call oDS_PH_PY313B.setValue("U_ColReg05", i, oRecordSet.Fields("PrRpmtDt").Value) '������ȯ����
        Call oDS_PH_PY313B.setValue("U_ColSum02", i, oRecordSet.Fields("PrRmainAmt").Value) '������ȯ�ܾ�
        Call oDS_PH_PY313B.setValue("U_ColReg06", i, oRecordSet.Fields("UseDt").Value) '����ϼ�
        Call oDS_PH_PY313B.setValue("U_ColSum03", i, oRecordSet.Fields("RpmtAmt").Value) '��ȯ�ݾ�
        Call oDS_PH_PY313B.setValue("U_ColSum04", i, oRecordSet.Fields("RpmtInt").Value) '��ȯ����
        Call oDS_PH_PY313B.setValue("U_ColSum05", i, oRecordSet.Fields("RmainAmt").Value) '��ȯ�ܾ�
        Call oDS_PH_PY313B.setValue("U_ColReg07", i, oRecordSet.Fields("RegYN").Value) '��Ͽ���
        Call oDS_PH_PY313B.setValue("U_ColReg08", i, oRecordSet.Fields("LineId").Value) '����ID
        Call oDS_PH_PY313B.setValue("U_ColReg09", i, oRecordSet.Fields("Cnt").Value) 'ȸ��
        
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
PH_PY313_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "W")
    Exit Sub
PH_PY313_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY313_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY313_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY313_Validate_Error
    PH_PY313_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY313A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY313_Validate = False
        GoTo PH_PY313_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY313_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY313_Validate_Error:
    PH_PY313_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY313_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY313_AddData() As Boolean
'******************************************************************************
'Function ID : PH_PY313_AddData()
'�ش��� : PH_PY313
'��� : ��αݰ�� ������ ����
'�μ� : ����
'��ȯ�� : True : ���� ����, False : ���� ����
'Ư�̻��� : ����
'******************************************************************************
On Error GoTo PH_PY313_AddData_Error

    PH_PY313_AddData = False
    
    Dim loopCount As Integer
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String '�����
    Dim CntcCode As String '���
    Dim LoanDoc As Integer '��αݹ�����ȣ
    Dim RpmtDate As String '��ȯ����
    Dim RpmtAmt As Double '��ȯ�ݾ�
    Dim RpmtInt As Double '��ȯ����
    Dim RmainAmt As Double '��ȯ�ܾ�
    Dim LineId As Integer '��αݶ��ι�ȣ
    Dim RegYN As String '��Ͽ���
    Dim UseDt As Integer '����� ��
    Dim Cnt As Integer 'ȸ��
    
    oMat1.FlushToDataSource
    For loopCount = 0 To oMat1.VisualRowCount - 1
        
        CLTCOD = oForm.Items("CLTCOD").Specific.Value '�����
        CntcCode = oDS_PH_PY313B.GetValue("U_ColReg02", loopCount) '���
        LoanDoc = oDS_PH_PY313B.GetValue("U_ColReg01", loopCount) '��αݹ�����ȣ
        RpmtDate = oForm.Items("RpmtDate").Specific.Value '��ȯ����
        RpmtAmt = oDS_PH_PY313B.GetValue("U_ColSum03", loopCount) '��ȯ�ݾ�
        RpmtInt = oDS_PH_PY313B.GetValue("U_ColSum04", loopCount) '��ȯ����
        RmainAmt = oDS_PH_PY313B.GetValue("U_ColSum05", loopCount) '��ȯ�ܾ�
        LineId = oDS_PH_PY313B.GetValue("U_ColReg08", loopCount) '��αݶ��ι�ȣ
        RegYN = oDS_PH_PY313B.GetValue("U_ColReg07", loopCount) '��Ͽ���
        UseDt = oDS_PH_PY313B.GetValue("U_ColReg06", loopCount) '����� ��
        Cnt = oDS_PH_PY313B.GetValue("U_ColReg09", loopCount) 'ȸ��
        
        sQry = "EXEC PH_PY313_02 '" & CLTCOD & "','" & CntcCode & "','" & LoanDoc & "','" & RpmtDate & "','" & RpmtAmt & "','" & RpmtInt & "','" & RmainAmt & "','" & LineId & "','" & RegYN & "','" & UseDt & "','" & Cnt & "'"
    
        Call oRecordSet.DoQuery(sQry)
        
    Next
    oMat1.LoadFromDataSource

    PH_PY313_AddData = True
    Set oRecordSet = Nothing
    
    Exit Function

PH_PY313_AddData_Error:
    PH_PY313_AddData = False
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY313_AddData_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

