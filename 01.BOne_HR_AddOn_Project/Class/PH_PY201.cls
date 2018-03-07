VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY201.cls
'//  Module         : �λ���� > �λ���������
'//  Desc           : �����ӹ��� �ް���� ���
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY201A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY201B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY201.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY201_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY201"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY201_CreateItems
    Call PH_PY201_EnableMenus
    Call PH_PY201_SetDocument(oFromDocEntry01)
'    Call PH_PY201_FormResize
    
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

Private Function PH_PY201_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY201_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY201A = oForm.DataSources.DBDataSources("@PH_PY201A")
    Set oDS_PH_PY201B = oForm.DataSources.DBDataSources("@PH_PY201B")

    
    Set oMat1 = oForm.Items("Mat01").Specific

    oMat1.SelectionMode = ms_Auto
    oMat1.AutoResizeColumns


    '//----------------------------------------------------------------------------------------------
    '// �⺻����
    '//----------------------------------------------------------------------------------------------
    
    '�����
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True

'    '�б�
'    Set oCombo = oForm.Items("Quarter").Specific
'    oCombo.ValidValues.Add "", ""
'    oCombo.ValidValues.Add "01", "1/4 Ȥ�� 1�б�"
'    oCombo.ValidValues.Add "02", "2/4"
'    oCombo.ValidValues.Add "03", "3/4 Ȥ�� 2�б�"
'    oCombo.ValidValues.Add "04", "4/4"
'    oCombo.Select 0, psk_Index
'    oForm.Items("Quarter").DisplayDesc = True
'
'    '��Ʈ����-����
'    Set oColumn = oMat1.Columns("Sex")
'    oColumn.ValidValues.Add "", ""
'    oColumn.ValidValues.Add "01", "����"
'    oColumn.ValidValues.Add "02", "����"
'    oColumn.DisplayDesc = True
'
'    '��Ʈ����-�б�
'    Set oColumn = oMat1.Columns("SchCls")
'    oColumn.ValidValues.Add "", ""
'    sQry = "            SELECT      T1.U_Code,"
'    sQry = sQry & "                 T1.U_CodeNm"
'    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
'    sQry = sQry & "                 INNER JOIN"
'    sQry = sQry & "                 [@PS_HR200L] AS T1"
'    sQry = sQry & "                     ON T0.Code = T1.Code"
'    sQry = sQry & "  WHERE      T0.Code = 'P222'"
'    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
'    sQry = sQry & "  ORDER BY  T1.U_Seq"
'
'    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
'
''    oColumn.ValidValues.Add "01", "�����б�"
''    oColumn.ValidValues.Add "02", "��������"
''    oColumn.ValidValues.Add "03", "���б�"
'    oColumn.DisplayDesc = True
'
'    '��Ʈ����-�г�
'    Set oColumn = oMat1.Columns("Grade")
'    oColumn.ValidValues.Add "", ""
'    oColumn.ValidValues.Add "01", "1�г�"
'    oColumn.ValidValues.Add "02", "2�г�"
'    oColumn.ValidValues.Add "03", "3�г�"
'    oColumn.ValidValues.Add "04", "4�г�"
'    oColumn.ValidValues.Add "05", "5�г�"
'    oColumn.DisplayDesc = True
'
'    '��Ʈ����-ȸ��
'    Set oColumn = oMat1.Columns("Count")
'    oColumn.ValidValues.Add "", ""
'    oColumn.ValidValues.Add "01", "1��"
'    oColumn.ValidValues.Add "02", "2��"
'    oColumn.DisplayDesc = True



    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY201_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY201_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY201_EnableMenus()
    
    On Error GoTo PH_PY201_EnableMenus_Error
    
    Call oForm.EnableMenu("1283", False)         '// ����
    Call oForm.EnableMenu("1287", True)          '// ����
'    Call oForm.EnableMenu("1286", True)         '// �ݱ�
    Call oForm.EnableMenu("1284", True)         '// ���
    Call oForm.EnableMenu("1293", True)         '// �����
    
    Exit Sub
    
PH_PY201_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY201_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY201_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY201_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY201_FormItemEnabled
        Call PH_PY201_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY201_FormItemEnabled
        oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY201_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY201_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY201_FormItemEnabled()
    On Error GoTo PH_PY201_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("DocDate").Enabled = True
        oForm.Items("Comment").Enabled = True
        
        '�� DocEntry ����
        Call PH_PY201_FormClear
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", False)    '//�����߰�
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("DocDate").Enabled = True
        oForm.Items("DocEntry").Enabled = True
        oForm.Items("Comment").Enabled = False
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("DocDate").Enabled = True
        oForm.Items("Comment").Enabled = True
       ' oMat1.Columns("MSTCOD").Editable = False
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//����ã��
        Call oForm.EnableMenu("1282", True)     '//�����߰�

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY201_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY201_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Dim SchCls As String '�б�(VALIDATE Event���� ���)
    Dim PayCnt As Integer '����Ƚ��(COMBO_SELECT Event���� ���)
    Dim Tuition As Double '��ϱݰ�(VALIDATE Event���� ���)
    Dim FeeTot As Double '���бݰ�(VALIDATE Event���� ���)
    Dim TuiTot As Double '��ϱݰ�(VALIDATE Event���� ���)
    Dim Total As Double '�Ѱ�(VALIDATE Event���� ���)
    
    Dim PreTuition As Double '��ϱ� �Է� �� ������
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY201_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        
                        '//�ؾ����� �۾�
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY201_DataValidCheck = False Then
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
                            Call PH_PY201_FormItemEnabled
                            Call PH_PY201_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY201_FormItemEnabled
                            Call PH_PY201_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY201_FormItemEnabled
                        End If
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                If pval.ItemUID = "Mat01" Then
        
                    If pval.ColUID = "MSTCOD" And pval.CharPressed = "9" Then
                    
                        If oMat1.Columns.Item("MSTCOD").Cells(pval.Row).Specific.VALUE = "" Then
                            Call Sbo_Application.ActivateMenuItem("7425")
                            BubbleEvent = False
                        End If
                    
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
'                    Call PH_PY201_AddMatrixRow
                    
                    If pval.ItemUID = "Mat01" Then
                        
                        Call oMat1.AutoResizeColumns
                    End If
                    
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
                            Call oDS_PH_PY201A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.VALUE & "'"))
                        
                        Case "Mat01"
                        
                            If pval.ColUID = "MSTCOD" Then
                            
                                oMat1.FlushToDataSource
                                
                                   Call oDS_PH_PY201B.setValue("U_MSTNAM", pval.Row - 1, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE & "'"))
                                   Call oDS_PH_PY201B.setValue("U_TeamCode", pval.Row - 1, MDC_GetData.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE & "'"))
                                   Call oDS_PH_PY201B.setValue("U_TeamName", pval.Row - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oDS_PH_PY201B.GetValue("U_TeamCode", pval.Row - 1) & "'", " AND Code = '1'"))
                                   Call oDS_PH_PY201B.setValue("U_RspCode", pval.Row - 1, MDC_GetData.Get_ReData("U_RspCode", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE & "'"))
                                   Call oDS_PH_PY201B.setValue("U_RspName", pval.Row - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oDS_PH_PY201B.GetValue("U_RspCode", pval.Row - 1) & "'", " AND Code = '2'"))
                                   Call oDS_PH_PY201B.setValue("U_ClsCode", pval.Row - 1, MDC_GetData.Get_ReData("U_ClsCode", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE & "'"))
                                   Call oDS_PH_PY201B.setValue("U_ClsName", pval.Row - 1, MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oDS_PH_PY201B.GetValue("U_ClsCode", pval.Row - 1) & "'", " AND Code = '9'"))
                                   Call oDS_PH_PY201B.setValue("U_Amount", pval.Row - 1, MDC_GetData.Get_ReData("TOP 1 U_Num1", "Code", "[@PS_HR200l]", "'P238' ORDER BY U_Num1 DESC"))
                                   
                                oMat1.LoadFromDataSource
                                
                                Call PH_PY201_AddMatrixRow
                            
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

                Call PH_PY201_FormItemEnabled
                Call PH_PY201_AddMatrixRow
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
                Set oDS_PH_PY201A = Nothing
                Set oDS_PH_PY201B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY201A", "Code")
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY201A", "DocEntry")      '//������ ���ѿ� ���� ����� ����
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY201_FormItemEnabled
                Call PH_PY201_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY201_FormItemEnabled
                Call PH_PY201_AddMatrixRow
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY201_FormItemEnabled
                Call PH_PY201_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY201_FormItemEnabled
            Case "1293" '// �����

                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY201B.Size - 1)
                        If oDS_PH_PY201B.GetValue("U_LineNum", i) = "" Then
                            oDS_PH_PY201B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY201B.Size
                        Call oDS_PH_PY201B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY201_AddMatrixRow
                
            '����
            Case "1287"
            
                Call oForm.Freeze(True)
                Call oDS_PH_PY201A.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat1.VisualRowCount - 1
                    Call oMat1.FlushToDataSource
                    Call oDS_PH_PY201B.setValue("DocEntry", i, "")
                    Call oDS_PH_PY201B.setValue("U_PayYN", i, "N")
                    Call oMat1.LoadFromDataSource
                Next i
                
                oForm.Items("Quarter").Enabled = True
                Call oForm.Freeze(False)
            
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

Sub PH_PY201_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY201_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY201B.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY201B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY201B.InsertRecord (oRow)
            End If
            oDS_PH_PY201B.Offset = oRow
            oDS_PH_PY201B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY201B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY201B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY201B.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY201B.setValue "U_TeamName", oRow, ""
            oDS_PH_PY201B.setValue "U_RspCode", oRow, ""
            oDS_PH_PY201B.setValue "U_RspName", oRow, ""
            oDS_PH_PY201B.setValue "U_ClsCode", oRow, ""
            oDS_PH_PY201B.setValue "U_ClsName", oRow, ""
            oDS_PH_PY201B.setValue "U_Amount", oRow, 0
            oDS_PH_PY201B.setValue "U_Comment", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY201B.Offset = oRow - 1
            oDS_PH_PY201B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY201B.setValue "U_MSTCOD", oRow - 1, ""
            oDS_PH_PY201B.setValue "U_MSTNAM", oRow - 1, ""
            oDS_PH_PY201B.setValue "U_TeamCode", oRow - 1, ""
            oDS_PH_PY201B.setValue "U_TeamName", oRow - 1, ""
            oDS_PH_PY201B.setValue "U_RspCode", oRow - 1, ""
            oDS_PH_PY201B.setValue "U_RspName", oRow - 1, ""
            oDS_PH_PY201B.setValue "U_ClsCode", oRow - 1, ""
            oDS_PH_PY201B.setValue "U_ClsName", oRow - 1, ""
            oDS_PH_PY201B.setValue "U_Amount", oRow - 1, 0
            oDS_PH_PY201B.setValue "U_Comment", oRow, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY201B.Offset = oRow
        oDS_PH_PY201B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY201B.setValue "U_MSTCOD", oRow, ""
        oDS_PH_PY201B.setValue "U_MSTNAM", oRow, ""
        oDS_PH_PY201B.setValue "U_TeamCode", oRow, ""
        oDS_PH_PY201B.setValue "U_TeamName", oRow, ""
        oDS_PH_PY201B.setValue "U_RspCode", oRow, ""
        oDS_PH_PY201B.setValue "U_RspName", oRow, ""
        oDS_PH_PY201B.setValue "U_ClsCode", oRow, ""
        oDS_PH_PY201B.setValue "U_ClsName", oRow, ""
        oDS_PH_PY201B.setValue "U_Amount", oRow, 0
        oDS_PH_PY201B.setValue "U_Comment", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY201_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY201_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY201_FormClear()
On Error GoTo PH_PY201_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY201'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY201_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY201_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY201_DataValidCheck() As Boolean
On Error GoTo PH_PY201_DataValidCheck_Error
    PH_PY201_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim CLTCOD As String
    Dim StdYear As String
    Dim Quarter As String
    Dim Count As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    
    
    '�����
    If Trim(oDS_PH_PY201A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "������� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY201_DataValidCheck = False
        Exit Function
    End If
    
    '��������
    If Trim(oDS_PH_PY201A.GetValue("U_DocDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "�������ڴ� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("DocDate").CLICK ct_Regular
        PH_PY201_DataValidCheck = False
        Exit Function
    End If
    
    
    '����
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
        
            '���
            If oMat1.Columns("MSTCOD").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "����� �ʼ��Դϴ�.", bmt_Short, True
                oMat1.Columns("MSTCOD").Cells(i).CLICK ct_Regular
                PH_PY201_DataValidCheck = False
                Exit Function
            End If
            
            '�ݾ�
            If oMat1.Columns("Amount").Cells(i).Specific.VALUE = "0" Then
                Sbo_Application.SetStatusBarMessage "�ݾ��� �ʼ��Դϴ�.", bmt_Short, True
                oMat1.Columns("Amount").Cells(i).CLICK ct_Regular
                PH_PY201_DataValidCheck = False
                Exit Function
            End If
            
'            '�ް���
'            If oMat1.Columns("VacLoc").Cells(i).Specific.VALUE = "" Then
'                Sbo_Application.SetStatusBarMessage "�ް����� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("VacLoc").Cells(i).CLICK ct_Regular
'                PH_PY201_DataValidCheck = False
'                Exit Function
'            End If
'
'            '�ް�������
'            If oMat1.Columns("VacFrDt").Cells(i).Specific.VALUE = "" Then
'                Sbo_Application.SetStatusBarMessage "�ް��������� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("VacFrDt").Cells(i).CLICK ct_Regular
'                PH_PY201_DataValidCheck = False
'                Exit Function
'            End If
'
'            '�ް�������
'            If oMat1.Columns("VacToDt").Cells(i).Specific.VALUE = "0" Then
'                Sbo_Application.SetStatusBarMessage "�ް��������� �ʼ��Դϴ�.", bmt_Short, True
'                oMat1.Columns("VacToDt").Cells(i).CLICK ct_Regular
'                PH_PY201_DataValidCheck = False
'                Exit Function
'            End If
            
        Next
    Else
        Sbo_Application.SetStatusBarMessage "���� �����Ͱ� �����ϴ�.", bmt_Short, True
        PH_PY201_DataValidCheck = False
        Exit Function
    End If
    

    
    
    
        oMat1.FlushToDataSource
    '// Matrix ������ �� ����(DB �����)
    If oDS_PH_PY201B.Size > 1 Then oDS_PH_PY201B.RemoveRecord (oDS_PH_PY201B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY201_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY201_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY201_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY201_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY201_MTX01()

    '//��Ʈ������ ������ �ε�
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY201_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("Param01").Specific.VALUE
    Param02 = oForm.Items("Param01").Specific.VALUE
    Param03 = oForm.Items("Param01").Specific.VALUE
    Param04 = oForm.Items("Param01").Specific.VALUE
    
    sQry = "SELECT 10"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "E")
        GoTo PH_PY201_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY201B.InsertRecord (i)
        End If
        oDS_PH_PY201B.Offset = i
        Call oDS_PH_PY201B.setValue("U_COL01", i, oRecordSet.Fields(0).VALUE)
        Call oDS_PH_PY201B.setValue("U_COL02", i, oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "�� ��ȸ��...!"
    Next
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY201_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY201_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY201_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY201_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY201_Validate_Error
    PH_PY201_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY201A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY201_Validate = False
        GoTo PH_PY201_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY201_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY201_Validate_Error:
    PH_PY201_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY201_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY201_CheckAmt(ByVal pAmt As Double, ByVal pSchCls As String) As Boolean
'******************************************************************************
'Function ID : PH_PY201_CheckAmt()
'�ش��� : PH_PY201
'��� : ��ϱ�(���ڱ�)�� �ѵ��ݾ� üũ
'�μ� : pAmt:�Է��� ���ڱ� �ݾ�, pSchCls:�б�����(�����б�:01, ��������:02, ���б�:03)
'��ȯ�� : �ѵ��ݾ� �ʰ� �ô� False�� ��ȯ, �׷��� ������ True
'Ư�̻��� : ����
'******************************************************************************
On Error GoTo PH_PY201_CheckAmt_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim CheckAmt As Double
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "           SELECT      U_Num1 AS [CheckAmt]"
    sQry = sQry & " FROM       [@PS_HR200L] AS T0"
    sQry = sQry & " WHERE      T0.Code = 'P222'"
    sQry = sQry & "                AND T0.U_Code = '" & pSchCls & "'"
    
    Call oRecordSet.DoQuery(sQry)
    
    CheckAmt = oRecordSet.Fields("CheckAmt").VALUE
    
    If CheckAmt < pAmt Then '�Է±ݾ��� �ѵ��ݾ׺��� ũ��
        
        PH_PY201_CheckAmt = False
    
    Else
    
        PH_PY201_CheckAmt = True
    
    End If
    
    Exit Function
PH_PY201_CheckAmt_Error:
    Set oRecordSet = Nothing
    PH_PY201_CheckAmt = False
    Sbo_Application.SetStatusBarMessage "PH_PY201_CheckAmt_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function





