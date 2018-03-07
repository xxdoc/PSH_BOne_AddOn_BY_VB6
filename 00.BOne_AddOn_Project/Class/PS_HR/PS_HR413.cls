VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR413"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_HR413.cls
'//  Module         : HR
'//  Description    : �������� ������ ���
'//  FormType       : PS_HR413
'//  Create Date    : 2013.06.10
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_HR413H  As SAPbouiCOM.DBDataSource '������
Private oDS_PS_HR413L  As SAPbouiCOM.DBDataSource '��϶���

Private oLastItemUID01 As String 'Ŭ�������� ������ ������ ������ Uid��
Private oLastColUID01  As String '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLastColRow01  As Long '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR413.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '��Ʈ������ Ÿ��Ʋ���̿� �����̸� ����
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_HR413_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "Code" '//UDO����϶�

    oForm01.EnableMenu ("1293"), True     '// �����
    oForm01.EnableMenu ("1287"), True     '// ����
    oForm01.EnableMenu ("1284"), True     '// ���
    oForm01.Freeze True
    
    Call CreateItems
    Call ComboBox_Setting
    
    Call SetDocument(oFromDocEntry01)
    'Call AddMatrixRow(0, True)
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                  
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                        oForm01.Items("Code").Specific.VALUE = Trim(oForm01.Items("BPLId").Specific.VALUE) + Trim(oForm01.Items("Year").Specific.VALUE) + Trim(oForm01.Items("RateCode").Specific.VALUE)

                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                  
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
                
                If pval.ItemUID = "Btn01" Then
'                    If oForm01.Items("RateCode").Specific.VALUE = "A30" Then
                        
                        Call DataLoad1
                        
'                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "Mat01" Then
                        If oMat01.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                        
                    ElseIf pval.ItemUID = "RateCode" Then
                        If oForm01.Items("RateCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then
                    If pval.Row > 0 Then
                        Call oMat01.SelectRow(pval.Row, True, False)
                    End If
                End If
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    FormItemEnabled
                    Call AddMatrixRow(0, True)
                End If
            If pval.BeforeAction = False And pval.ItemChanged = True Then
               If pval.ColUID = "MSTCOD" Then
                  FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
               End If
            End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If (pval.ItemUID = "Mat01") Then
                        If pval.ColUID = "MSTCOD" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    Else
                        If (pval.ItemUID = "RateCode") Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                Call AddMatrixRow(oMat01.VisualRowCount, False)
                Call FormItemEnabled
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error

    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim sSeq$

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
        
' Matrix �ʵ忡 ���� ���� â ����ֱ�
    If oUID = "Mat01" Then
        Select Case oCol
            Case "MSTCOD"
                
                '--------------------------------------------------------------------------------------------
                If oRow = oMat01.RowCount And Trim(oDS_PS_HR413L.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
                '// ���� ���� �߰�
                    Call AddMatrixRow(0, False)
                    oMat01.Columns("MSTCOD").Cells(oRow).Click ct_Regular
                End If
                
                oMat01.FlushToDataSource
                
                sQry = "        Select  FULLNAME = t.U_FULLNAME "
                sQry = sQry & " From    [@PH_PY001A] t "
                sQry = sQry & " Where   Code =  '" & oMat01.Columns("MSTCOD").Cells(oRow).Specific.VALUE & "' "
                sQry = sQry & "         And t.U_CLTCOD = '" & oForm01.Items("BPLId").Specific.VALUE & "'"
                
                oRecordSet01.DoQuery sQry
                oDS_PS_HR413L.setValue "U_FULLNAME", oRow - 1, Trim(oRecordSet01.Fields(0).VALUE)
                
                oMat01.LoadFromDataSource
       End Select
    ElseIf oUID = "RateCode" Then
    
        sQry = "        Select  U_RateMNm + Case When U_RateSNm = '' Then '' Else (Case When Isnull(U_RateSNm,'') = '' Then '' Else '-' + Isnull(U_RateSNm,'') End) End "
        sQry = sQry & " From    [@PS_HR400H] a "
        sQry = sQry & "         inner Join "
        sQry = sQry & "         [@PS_HR400L] b "
        sQry = sQry & "             On a.Code = b.Code "
        sQry = sQry & " Where   a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "'"
        sQry = sQry & "         and U_Year = '" & Trim(oForm01.Items("Year").Specific.VALUE) & "'"
        sQry = sQry & "         And b.U_RateCode = '" & Trim(oForm01.Items("RateCode").Specific.VALUE) & "'"

        oRecordSet01.DoQuery sQry
        oForm01.Items("RateMNm").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    End If
   
    Set oRecordSet01 = Nothing
    
    Exit Sub
       
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�ڵ带 Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo MenuEvent_Error
    
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                  Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": 'ã��
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO����϶�
            Case "1282": '�߰�
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO����϶�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource        ' DBDataSource�� ���ڵ尡 ���� �� �����.
                    Call oDS_PS_HR413L.RemoveRecord(oDS_PS_HR413L.Size - 1)     ' ���ڵ� �� ���� �����.
                    oMat01.LoadFromDataSource       ' DBDataSource�� ��Ʈ������ �ø���
                    If oMat01.RowCount = 0 Then
'                        Call AddMatrixRow(1, 0, True)
                        Call AddMatrixRow(1, 0)
                    Else
                        If Trim(oDS_PS_HR413L.GetValue("U_MSTCOD", oMat01.RowCount - 1)) <> "" Then
'                            Call AddMatrixRow(1, oMat01.RowCount, True)
                            Call AddMatrixRow(1, oMat01.RowCount)
                        End If
                    End If
                End If
            Case "1281": 'ã��
                 Call AddMatrixRow(0, True) '//UDO���
                 Call FormItemEnabled
            Case "1287": '����
                oForm01.Freeze True
                oDS_PS_HR413H.setValue "Code", 0, ""
                oDS_PS_HR413H.setValue "U_Year", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_HR413L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
            Case "1282": '�߰�
                 Call FormItemEnabled '//UDO���
                 Call AddMatrixRow(0, True) '//UDO���
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư             '�߰���ư Ŭ���� Matrix�� �� �߰�
                 Call FormItemEnabled
                 Call AddMatrixRow(1, oMat01.RowCount)
        End Select
    End If
    
    Exit Sub
    
MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function CreateItems() As Boolean
On Error GoTo CreateItems_Error

    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
       
    Set oDS_PS_HR413H = oForm01.DataSources.DBDataSources("@PS_HR413H")
    Set oDS_PS_HR413L = oForm01.DataSources.DBDataSources("@PS_HR413L")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    
'    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
'    oForm01.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
'    oForm01.DataSources.UserDataSources.Item("DocDate").Value = Format(Now, "YYYYMMDD")
    
    oForm01.Items("Year").Specific.VALUE = Left(Date, 4)
'    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
   
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
    
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error

    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
             
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//���̵� ����� ����
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Call oForm01.Freeze(False)
    
    Exit Sub
    
ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error

    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
        'Call FormClear '//UDO���
        Call oForm01.EnableMenu("1281", True) '//ã��
        Call oForm01.EnableMenu("1282", False) '//�߰�
        oForm01.Items("Code").Enabled = False
        oForm01.Items("Year").Enabled = True
        oForm01.Items("RateCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
        Call oForm01.EnableMenu("1281", True) '//ã��
        oForm01.Items("Code").Enabled = False
        oForm01.Items("Year").Enabled = True
        oForm01.Items("RateCode").Enabled = True
        Call oForm01.EnableMenu("1282", True) '//�߰�
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("Code").Enabled = False
        oForm01.Items("Year").Enabled = False
        oForm01.Items("RateCode").Enabled = False
        Call oForm01.EnableMenu("1282", True) '//�߰�
      
        '//����忡���� �����ۼ���
    End If
    Call oForm01.Freeze(False)
    
    Exit Sub
    
FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo AddMatrixRow_Error
    
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//���߰�����
        oRow = oMat01.RowCount
        oDS_PS_HR413L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_HR413L.Offset = oRow
    oDS_PS_HR413L.setValue "LineId", oRow, oRow + 1
    oDS_PS_HR413L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    
    Exit Sub
    
AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//�۾�
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    
    Exit Sub
    
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error

    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//������� ��������ɿ��ΰ˻�
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineId").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_HR413L.RemoveRecord(oDS_PS_HR413L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call AddMatrixRow(0)
            Else
                If Trim(oDS_PS_HR413L.GetValue("U_MSTCOD", oMat01.RowCount - 1)) <> "" Then
                    Call AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    
    Exit Sub
    
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error

    Dim i             As Long
    Dim ErrNum        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    Dim sQry          As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    ErrNum = 0
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// ȭ����� ��Ʈ������ �Էµ� ������ ��� ������ͼҽ��� �ѱ��
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.FlushToDataSource
    
    '// ����
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 Then
        If oDS_PS_HR413L.GetValue("U_MSTCOD", 0) = "" Then
            ErrNum = 1
            GoTo MatrixSpaceLineDel_Error
        End If
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// �Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '// �̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 2
            oDS_PS_HR413L.Offset = i
            Select Case True
                Case oDS_PS_HR413L.GetValue("U_MSTCOD", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
                Case oDS_PS_HR413L.GetValue("U_Qty", i) = 0
                    ErrNum = 3
                    GoTo MatrixSpaceLineDel_Error
                Case oDS_PS_HR413L.GetValue("U_Value", i) = 0
                    ErrNum = 4
                    GoTo MatrixSpaceLineDel_Error
            End Select
        Next i
        
        If oDS_PS_HR413L.GetValue("U_MSTCOD", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_HR413L.RemoveRecord oMat01.VisualRowCount - 1
        End If
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���� �����Ͽ����� DB������ �ҽ��� �ٽ� �����´�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
    
    Exit Function
    
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���ε���Ÿ�� �����ϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "����� �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "������ �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "������ �ʼ��Է»����Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error

    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_HR413H.GetValue("U_Year", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_HR413H.GetValue("U_RateCode", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select
   
   HeaderSpaceLineDel = True
   
    Exit Function

HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�⵵�� �ʼ��Է� �����Դϴ�.", "E"
    End If
    If ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "���׸��� �ʼ��Է� �����Դϴ�.", "E"
    End If
    If ErrNum = "" Then
        MDC_Com.MDC_GF_Message "������ �Ǿ����ϴ�.", "W"
    End If
    HeaderSpaceLineDel = False
End Function

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error

    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then '//BeforeAction = False
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    End If
    
    Exit Sub
    
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo SetDocument_Error

    If (oFromDocEntry01 = "") Then
        Call FormItemEnabled
        Call AddMatrixRow(0, True)
    Else
        oForm01.Mode = fm_FIND_MODE
        Call FormItemEnabled
        oForm01.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    
    Exit Sub
    
SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub DataLoad1()
On Error GoTo DataLoad1_Error

    Dim ErrNum      As Integer
    Dim DocNum      As String
    Dim BPLID       As String
    Dim YmFrom      As String
    Dim YmTo        As String
    Dim RateCode    As String '���׸�
    Dim Query01     As String
    Dim sRow        As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    ErrNum = 0
    
    Call oForm01.Freeze(True)
    
    sRow = 1
    '// Check
    Select Case True
        Case oDS_PS_HR413H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo DataLoad1_Error
        Case oDS_PS_HR413H.GetValue("U_Ymfrom", 0) = ""
            ErrNum = 2
            GoTo DataLoad1_Error
        Case oDS_PS_HR413H.GetValue("U_YmTo", 0) = ""
            ErrNum = 3
            GoTo DataLoad1_Error
        Case oDS_PS_HR413H.GetValue("U_RateCode", 0) = ""
            ErrNum = 4
            GoTo DataLoad1_Error
    End Select
   
   BPLID = oDS_PS_HR413H.GetValue("U_BPLId", 0)
   YmFrom = oDS_PS_HR413H.GetValue("U_YmFrom", 0)
   YmTo = oDS_PS_HR413H.GetValue("U_YmTo", 0)
   RateCode = oDS_PS_HR413H.GetValue("U_RateCode", 0)
   
   Query01 = "         EXEC PS_HR413_01 '"
   Query01 = Query01 & BPLID & "','"
   Query01 = Query01 & RateCode & "','"
   Query01 = Query01 & YmFrom & "','"
   Query01 = Query01 & YmTo & "'"
   
'   Query01 = "           Select     g.sabun, "
'   Query01 = Query01 & "            g.kname, "
'   Query01 = Query01 & "            g.cnt, "
'   Query01 = Query01 & "            jumsu = Case When g.a + g.b > 5 Then 5 Else g.a + g.b End,"
'   Query01 = Query01 & "            bigo =  Case "
'   Query01 = Query01 & "                        When g.A1 > 0 And g.A2 = 0 Then '����' + Convert(Nvarchar(5), g.A1)"
'   Query01 = Query01 & "                        When g.A1 > 0 And g.A2 > 0 Then '����' + Convert(Nvarchar(5), g.A1) + ',���' + Convert(Nvarchar(5), g.A2)"
'   Query01 = Query01 & "                        When g.A1 = 0 And g.A2 > 0 Then '���' + Convert(Nvarchar(5), g.A2) "
'   Query01 = Query01 & "                    End"
'   Query01 = Query01 & " From       ("
'   Query01 = Query01 & "                SELECT   a.sabun,"
'   Query01 = Query01 & "                         a.kname,"
'   Query01 = Query01 & "                         A1 = Sum(CASE WHEN a.div = '0' THEN 1 Else 0 End),"
'   Query01 = Query01 & "                         A2 = Sum(CASE WHEN a.div = '0' THEN 0 Else 1 End),"
'   Query01 = Query01 & "                         cnt = Sum(CASE WHEN a.div = '0' THEN 1 Else 0 End) + Sum(CASE WHEN a.div = '0' THEN 0 Else 1 End),"
'   Query01 = Query01 & "                         a = Case When Sum(CASE WHEN a.div = '0' THEN 1 Else 0 End) > 0 Then 5 Else 0 End,"
'   Query01 = Query01 & "                         b =  Case "
'   Query01 = Query01 & "                                    When Sum(CASE WHEN a.div = '0' THEN 0 Else 1 End) / 6.0 * 100 >= 120 Then 5"
'   Query01 = Query01 & "                                    When Sum(CASE WHEN a.div = '0' THEN 0 Else 1 End) / 6.0 * 100 >= 100 Then 4"
'   Query01 = Query01 & "                                    When Sum(CASE WHEN a.div = '0' THEN 0 Else 1 End) / 6.0 * 100 >= 80 Then 3"
'   Query01 = Query01 & "                                    When Sum(CASE WHEN a.div = '0' THEN 0 Else 1 End) / 6.0 * 100 >= 60 Then 2 "
'   Query01 = Query01 & "                                    Else 1 "
'   Query01 = Query01 & "                              End"
'   Query01 = Query01 & "                FROM     [ZPS_QM151] a"
'   Query01 = Query01 & "                WHERE    a.saup = '" & BPLId & "'"
'   Query01 = Query01 & "                         and Convert(char(6),a.ymd,112) between '" & YmFrom & "' and '" & YmTo & "'"
'   Query01 = Query01 & "                         and Isnull(a.par_a,0) <> 0"
'   Query01 = Query01 & "                         and a.div <> '2'"
'   Query01 = Query01 & "                Group by a.sabun, "
'   Query01 = Query01 & "                         a.kname "
'   Query01 = Query01 & "            ) g"
   
   Call RecordSet01.DoQuery(Query01)
   
   Call oMat01.Clear
   Call oDS_PS_HR413L.Clear
   Call oMat01.FlushToDataSource
   Call oMat01.LoadFromDataSource
   
   Do Until RecordSet01.EOF
        
        Call oDS_PS_HR413L.setValue("U_LineNum", sRow - 1, sRow)
        Call oDS_PS_HR413L.setValue("U_MSTCOD", sRow - 1, Trim(RecordSet01.Fields(0).VALUE)) '���
        Call oDS_PS_HR413L.setValue("U_FULLNAME", sRow - 1, Trim(RecordSet01.Fields(1).VALUE)) '����
        Call oDS_PS_HR413L.setValue("U_Qty", sRow - 1, Trim(RecordSet01.Fields(2).VALUE)) '�Ǽ�
        Call oDS_PS_HR413L.setValue("U_Value", sRow - 1, Trim(RecordSet01.Fields(3).VALUE)) '����
        Call oDS_PS_HR413L.setValue("U_Comments", sRow - 1, Trim(RecordSet01.Fields(4).VALUE)) '���
        
        Call AddMatrixRow(sRow, False)
        sRow = sRow + 1
        RecordSet01.MoveNext
   Loop
   
   Call oMat01.LoadFromDataSource
   Call oMat01.AutoResizeColumns
   
   Set RecordSet01 = Nothing
   Call oForm01.Freeze(False)
   
    Exit Sub

DataLoad1_Error:

    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "������� �ʼ��Է� �����Դϴ�.", "E"
        Call oForm01.Items("BPLId").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "���س��(����)�� �ʼ��Է� �����Դϴ�.", "E"
        Call oForm01.Items("YmFrom").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "���س��(����)�� �ʼ��Է� �����Դϴ�.", "E"
        Call oForm01.Items("YmTo").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "���׸��� �ʼ��Է� �����Դϴ�.", "E"
        Call oForm01.Items("RateCode").Click(ct_Regular)
    ElseIf ErrNum = "" Then
        MDC_Com.MDC_GF_Message "������ �Ǿ����ϴ�.", "W"
    End If

End Sub