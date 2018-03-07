VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY417"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY417.cls
'//  Module         : �޿����� > �������
'//  Desc           : �������� �������ϻ���
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY417A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY417B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY417.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY417_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY417"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY417_CreateItems
    Call PH_PY417_EnableMenus
    Call PH_PY417_SetDocument(oFromDocEntry01)
'    Call PH_PY417_FormResize
    
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

Private Function PH_PY417_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY417_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY417A = oForm.DataSources.DBDataSources("@PH_PY417A")
    Set oDS_PH_PY417B = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
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
        
    '���س��_S
    Call oForm.DataSources.UserDataSources.Add("YYYYMM", dt_SHORT_TEXT, 6)
    Call oForm.Items("YYYYMM").Specific.DataBind.SetBound(True, "", "YYYYMM")
    '���س��_E
    
    '��������_S
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
    Call oForm.Items("DocDate").Specific.DataBind.SetBound(True, "", "DocDate")
    '��������_E
    
    
    'ȯ��/¡������_S
    Call oForm.DataSources.UserDataSources.Add("Div", dt_SHORT_TEXT, 5)
    Call oForm.Items("Div").Specific.DataBind.SetBound(True, "", "Div")
    'ȯ��/¡������_E
    
    
    '�ҵ漼��_S
    Call oForm.DataSources.UserDataSources.Add("STot", dt_SUM, 50)
    Call oForm.Items("STot").Specific.DataBind.SetBound(True, "", "STot")
    '�ҵ漼��_E
    
    '�ֹμ���_S
    Call oForm.DataSources.UserDataSources.Add("JTot", dt_SUM, 50)
    Call oForm.Items("JTot").Specific.DataBind.SetBound(True, "", "JTot")
    '�ֹμ���_E
    
    '��Ư����_S
    Call oForm.DataSources.UserDataSources.Add("NTot", dt_SUM, 50)
    Call oForm.Items("NTot").Specific.DataBind.SetBound(True, "", "NTot")
    '��Ư����_E
    
    '�Ѱ�_S
    Call oForm.DataSources.UserDataSources.Add("Total", dt_SUM, 50)
    Call oForm.Items("Total").Specific.DataBind.SetBound(True, "", "Total")
    '�Ѱ�_E

    
    '���޳��_S
    Call oForm.DataSources.UserDataSources.Add("YM", dt_SHORT_TEXT, 6)
    Call oForm.Items("YM").Specific.DataBind.SetBound(True, "", "YM")
    '���޳��_E
        
    '��������_S
    Call oForm.DataSources.UserDataSources.Add("JOBTYP", dt_SHORT_TEXT, 5)
    Call oForm.Items("JOBTYP").Specific.DataBind.SetBound(True, "", "JOBTYP")
    '��������_E
    
    '���ޱ���_S
    Call oForm.DataSources.UserDataSources.Add("JOBGBN", dt_SHORT_TEXT, 5)
    Call oForm.Items("JOBGBN").Specific.DataBind.SetBound(True, "", "JOBGBN")
    '���ޱ���_E

    Dim oChkBox As SAPbouiCOM.CheckBox

    Call oForm.DataSources.UserDataSources.Add("Check01", dt_SHORT_TEXT, 1)
    Set oChkBox = oForm.Items("Check01").Specific
    oChkBox.ValOn = "Y"
    oChkBox.ValOff = "N"
    oChkBox.DataBind.SetBound True, , "Check01"
    oForm.DataSources.UserDataSources("Check01").VALUE = "N" '��üũ�� ���� �ְ� ���� �ε�
    
    '//----------------------------------------------------------------------------------------------
    '// �⺻����
    '//----------------------------------------------------------------------------------------------
    
    '�����
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '�б�
    Set oCombo = oForm.Items("Div").Specific
    oCombo.ValidValues.Add "00", "��ü"
    oCombo.ValidValues.Add "01", "ȯ��"
    oCombo.ValidValues.Add "02", "¡��"
    
    oCombo.Select 0, psk_Index
    oForm.Items("Div").DisplayDesc = True
    
    
    '// ��������
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.ValidValues.Add "1", "�޿�"
    'oCombo.ValidValues.Add "2", "��"
    oForm.Items("JOBTYP").DisplayDesc = True
    
    '// ���ޱ���
    Set oCombo = oForm.Items("JOBGBN").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JOBGBN").DisplayDesc = True
    
    oForm.Items("BtnPay").Enabled = False
    oForm.Items("YM").Enabled = False
    oForm.Items("JOBTYP").Enabled = False
    oForm.Items("JOBGBN").Enabled = False
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY417_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY417_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY417_EnableMenus()
    
    On Error GoTo PH_PY417_EnableMenus_Error
    
'    Call oForm.EnableMenu("1283", False)         '// ����
'    Call oForm.EnableMenu("1287", True)          '// ����
''    Call oForm.EnableMenu("1286", True)         '// �ݱ�
'    Call oForm.EnableMenu("1284", True)         '// ���
'    Call oForm.EnableMenu("1293", True)         '// �����
    
    Exit Sub
    
PH_PY417_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY417_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY417_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY417_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY417_FormItemEnabled
'        Call PH_PY417_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY417_FormItemEnabled
        oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY417_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY417_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY417_FormItemEnabled()
    On Error GoTo PH_PY417_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim CLTCOD As String
        
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        '�� DocEntry ����
'        Call PH_PY417_FormClear
        
        '// �����ڿ� ���� ���Ѻ� ����� �޺��ڽ�����
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '�⵵ ����
'        Call oDS_PH_PY417A.setValue("U_StdYear", 0, Format(Date, "YYYY"))
        oForm.Items("StdYear").Specific.VALUE = Format(Date, "YYYY") - 1

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
    
PH_PY417_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY417_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry As String
    Dim i As Long
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Dim StdYear As String '�⵵
    Dim CLTCOD As String '�����
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If PH_PY417_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'
'                        '//�ؾ����� �۾�
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If PH_PY417_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'                        '//�ؾ����� �۾�
'
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                    End If
'                End If
                If pval.ItemUID = "BtnSearch" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY417_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'
'                        '//�ؾ����� �۾�
                        Call PH_PY417_MTX01
                        
                    End If
                ElseIf pval.ItemUID = "BtnPay" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY417_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        Call PH_PY417_MTX01
                        
                        If Val(oForm.Items("Total").Specific.VALUE) <> 0 Then
                            Call PH_PY417_PY109_Update
                        Else
                            Sbo_Application.SetStatusBarMessage "�޿������ڷῡ ������ �������� ¡���ڷᰡ �����ϴ�.", bmt_Short, True
                        End If
                        
                    End If
                End If
                If pval.ItemUID = "Check01" Then
                    
                    If oForm.DataSources.UserDataSources("Check01").VALUE = "Y" Then
                        oForm.Items("YM").Enabled = True
                        oForm.Items("BtnPay").Enabled = True
                        oForm.Items("JOBTYP").Enabled = True
                        oForm.Items("JOBGBN").Enabled = True
                        oForm.Items("YM").CLICK ct_Regular
                    Else
                        oForm.Items("StdYear").CLICK ct_Regular
                        oForm.Items("YM").Enabled = False
                        oForm.Items("BtnPay").Enabled = False
                        oForm.Items("JOBTYP").Enabled = False
                        oForm.Items("JOBGBN").Enabled = False
                    End If
                End If
                    
                
            ElseIf pval.BeforeAction = False Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY417_FormItemEnabled
'                            Call PH_PY417_AddMatrixRow
'                        End If
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY417_FormItemEnabled
'                            Call PH_PY417_AddMatrixRow
'                        End If
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY417_FormItemEnabled
'                        End If
'                    End If
'                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                If pval.ItemUID = "Mat01" Then
        
                    
                
                ElseIf pval.ItemUID = "CntcCode" And pval.CharPressed = "9" Then
                
            
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
'                    Call PH_PY417_AddMatrixRow
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
                        Case "StdYear"
                            CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
                            StdYear = oForm.Items("StdYear").Specific.VALUE
                            
                            '//�ش�⵵�� ������ �޿������ ��������
                            sQry = "SELECT Distinct YM = U_YM, JIGBIL = Convert(char(8),U_JIGBIL,112) FROM [@PH_PY112A] WHERE U_JOBTYP = '1' And U_JOBGBN = '1' And U_CLTCOD =  '" & CLTCOD & "' And U_YM =  '" & StdYear & "12' "
                            oRecordSet.DoQuery sQry
                            oForm.Items("YYYYMM").Specific.VALUE = Trim(oRecordSet.Fields("YM").VALUE)
                            oForm.Items("DocDate").Specific.VALUE = Trim(oRecordSet.Fields("JIGBIL").VALUE)
                        
                    End Select

                End If
                
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY417_FormItemEnabled
                Call PH_PY417_AddMatrixRow
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
                Set oDS_PH_PY417A = Nothing
                Set oDS_PH_PY417B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY417A", "Code")
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
                Call PH_PY417_AddMatrixRow
            
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY417_FormItemEnabled
                Call PH_PY417_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//����ã��
                Call PH_PY417_FormItemEnabled
                Call PH_PY417_AddMatrixRow
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//�����߰�
                Call PH_PY417_FormItemEnabled
                Call PH_PY417_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY417_FormItemEnabled
            Case "1293" '// �����

                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY417B.Size - 1)
                        If oDS_PH_PY417B.GetValue("U_LineNum", i) = "" Then
                            oDS_PH_PY417B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY417B.Size
                        Call oDS_PH_PY417B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY417_AddMatrixRow
                
            Case "7169": '���� ��������
            
                '���� �������� ���� ó��
                Call oForm.Freeze(True)
                Call oDS_PH_PY417B.RemoveRecord(oDS_PH_PY417B.Size - 1)
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

Sub PH_PY417_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY417_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY417B.GetValue("U_LineNum", oRow - 1)) <> "" Then
            If oDS_PH_PY417B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY417B.InsertRecord (oRow)
            End If
            oDS_PH_PY417B.Offset = oRow
            oDS_PH_PY417B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY417B.setValue "U_ColReg01", oRow, ""
            oDS_PH_PY417B.setValue "U_ColReg02", oRow, ""
            oDS_PH_PY417B.setValue "U_ColReg03", oRow, ""
            oDS_PH_PY417B.setValue "U_ColSum01", oRow, ""
            oDS_PH_PY417B.setValue "U_ColSum02", oRow, ""
            oDS_PH_PY417B.setValue "U_ColSum03", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY417B.Offset = oRow - 1
            oDS_PH_PY417B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY417B.setValue "U_ColReg01", oRow - 1, ""
            oDS_PH_PY417B.setValue "U_ColReg02", oRow - 1, ""
            oDS_PH_PY417B.setValue "U_ColReg03", oRow - 1, ""
            oDS_PH_PY417B.setValue "U_ColSum01", oRow - 1, ""
            oDS_PH_PY417B.setValue "U_ColSum02", oRow - 1, ""
            oDS_PH_PY417B.setValue "U_ColSum03", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY417B.Offset = oRow
        oDS_PH_PY417B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY417B.setValue "U_ColReg01", oRow, ""
        oDS_PH_PY417B.setValue "U_ColReg02", oRow, ""
        oDS_PH_PY417B.setValue "U_ColReg03", oRow, ""
        oDS_PH_PY417B.setValue "U_ColSum01", oRow, ""
        oDS_PH_PY417B.setValue "U_ColSum02", oRow, ""
        oDS_PH_PY417B.setValue "U_ColSum03", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY417_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY417_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY417_FormClear()
On Error GoTo PH_PY417_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY417'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY417_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY417_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY417_DataValidCheck() As Boolean
On Error GoTo PH_PY417_DataValidCheck_Error
    PH_PY417_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '�����
    If Trim(oForm.Items("CLTCOD").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "������� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY417_DataValidCheck = False
        Exit Function
    End If
    
    '�⵵
    If Trim(oForm.Items("StdYear").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "�⵵�� �ʼ��Դϴ�.", bmt_Short, True
        oForm.Items("StdYear").CLICK ct_Regular
        PH_PY417_DataValidCheck = False
        Exit Function
    End If
    
    
    If oForm.DataSources.UserDataSources("Check01").VALUE = "Y" Then
        If Trim(oForm.Items("YM").Specific.VALUE) = "" Then
            Sbo_Application.SetStatusBarMessage "���޳���� �ʼ��Դϴ�.", bmt_Short, True
            oForm.Items("YM").CLICK ct_Regular
            PH_PY417_DataValidCheck = False
            Exit Function
        End If
        If Trim(oForm.Items("JOBTYP").Specific.VALUE) = "" Then
            Sbo_Application.SetStatusBarMessage "���������� �ʼ��Դϴ�.", bmt_Short, True
            oForm.Items("JOBTYP").CLICK ct_Regular
            PH_PY417_DataValidCheck = False
            Exit Function
        End If
        If Trim(oForm.Items("JOBGBN").Specific.VALUE) = "" Then
            Sbo_Application.SetStatusBarMessage "���ޱ����� �ʼ��Դϴ�.", bmt_Short, True
            oForm.Items("JOBGBN").CLICK ct_Regular
            PH_PY417_DataValidCheck = False
            Exit Function
        End If
    End If

    
        oMat1.FlushToDataSource
    '// Matrix ������ �� ����(DB �����)
    If oDS_PH_PY417B.Size > 1 Then oDS_PH_PY417B.RemoveRecord (oDS_PH_PY417B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY417_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY417_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY417_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY417_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY417_MTX01()

    '//��Ʈ������ ������ �ε�
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    
    
    Dim STot As Double '�ҵ漼��
    Dim JTot As Double '�ֹμ���
    Dim NTot As Double '��Ư����
    Dim Tot As Double '�Ѱ�
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY417_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oForm.Items("StdYear").Specific.VALUE
    Param03 = oForm.Items("YYYYMM").Specific.VALUE
    Param04 = oForm.Items("DocDate").Specific.VALUE
    Param05 = oForm.Items("Div").Specific.VALUE
    
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet.RecordCount, False)
    
    sQry = "EXEC PH_PY417_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call oMat1.Clear
        GoTo PH_PY417_MTX01_Exit
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY417B.InsertRecord (i)
        End If
        oDS_PH_PY417B.Offset = i
        Call oDS_PH_PY417B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY417B.setValue("U_ColReg05", i, oRecordSet.Fields("Div").VALUE)
        Call oDS_PH_PY417B.setValue("U_ColReg04", i, oRecordSet.Fields("BankCode").VALUE)
        Call oDS_PH_PY417B.setValue("U_ColReg01", i, oRecordSet.Fields("BankName").VALUE)
        Call oDS_PH_PY417B.setValue("U_ColReg02", i, oRecordSet.Fields("CntcName").VALUE)
        Call oDS_PH_PY417B.setValue("U_ColReg03", i, oRecordSet.Fields("AcctNo").VALUE)
        Call oDS_PH_PY417B.setValue("U_ColSum01", i, oRecordSet.Fields("SAmt").VALUE)
        Call oDS_PH_PY417B.setValue("U_ColSum02", i, oRecordSet.Fields("JAmt").VALUE)
        Call oDS_PH_PY417B.setValue("U_ColSum03", i, oRecordSet.Fields("NAmt").VALUE)
        Call oDS_PH_PY417B.setValue("U_ColSum04", i, oRecordSet.Fields("JSAmt").VALUE)
        
        STot = STot + oRecordSet.Fields("SAmt").VALUE
        JTot = JTot + oRecordSet.Fields("JAmt").VALUE
        NTot = NTot + oRecordSet.Fields("NAmt").VALUE
        Tot = Tot + oRecordSet.Fields("JSAmt").VALUE
        
        oRecordSet.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "�� ��ȸ��...!"
        
    Next
    
    oForm.Items("STot").Specific.VALUE = STot
    oForm.Items("JTot").Specific.VALUE = JTot
    oForm.Items("NTot").Specific.VALUE = NTot
    oForm.Items("Total").Specific.VALUE = Tot
    

    
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY417_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "W")
    Exit Sub
PH_PY417_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY417_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY417_PY109_Update()


    '//�޿������ڷῡ Update
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim Param05 As String
    Dim Param06 As String
    
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY417_PY109_Update_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oForm.Items("StdYear").Specific.VALUE
    Param03 = oForm.Items("YM").Specific.VALUE
    
    Param04 = oForm.Items("DocDate").Specific.VALUE
    Param05 = oForm.Items("JOBTYP").Specific.VALUE
    Param06 = oForm.Items("JOBGBN").Specific.VALUE
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet.RecordCount, False)
    
    sQry = "EXEC PH_PY417_02 '" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
'        Call MDC_Com.MDC_GF_Message("�޿������ڷῡ ���ڱ��ڷḦ ���ε� �����߽��ϴ�.", "W")
        Sbo_Application.SetStatusBarMessage "�޿������ڷῡ �������� ¡���ڷḦ ���ε� �����߽��ϴ�.", bmt_Short, True
    Else
        If Val(oRecordSet.Fields(0).VALUE) = -1 Then
            Sbo_Application.MessageBox ("�޿������ڷᰡ �����ϴ�. Ȯ�ιٶ��ϴ�")
        ElseIf Val(oRecordSet.Fields(0).VALUE) = 0 Then
            Sbo_Application.MessageBox ("�޿������ڷῡ �������� ¡���ڷḦ ���ε� ���� ���߽��ϴ�.Ȯ�ιٶ��ϴ�")
        ElseIf Val(oRecordSet.Fields(0).VALUE) > 0 Then
            Sbo_Application.MessageBox ("�������� �������� ¡���ڷḦ ���ε� �߽��ϴ�. �޿������ڷḦ Ȯ���ϼ���")
        End If
    End If
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY417_PY109_Update_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Call MDC_Com.MDC_GF_Message("����� �������� �ʽ��ϴ�.", "W")
    Exit Sub
PH_PY417_PY109_Update_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY417_PY109_Update_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY417_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY417_Validate_Error
    PH_PY417_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY417A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "�ش繮���� �ٸ�����ڿ� ���� ��ҵǾ����ϴ�. �۾��� �����Ҽ� �����ϴ�.", bmt_Short, True
        PH_PY417_Validate = False
        GoTo PH_PY417_Validate_Exit
    End If
'
    If ValidateType = "����" Then

    ElseIf ValidateType = "�����" Then

    ElseIf ValidateType = "���" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY417_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY417_Validate_Error:
    PH_PY417_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY417_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function