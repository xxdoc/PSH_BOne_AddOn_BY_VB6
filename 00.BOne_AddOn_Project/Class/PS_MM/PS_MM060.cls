VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM060.cls
'//  Module         : SY
'//  Desc           : ���� ǰ��/�԰ݺ� �ܰ����
'//  FormType       : PS_MM060
'//  Create Date    : 2010.09.15
'//  Modified Date  :
'//  Creator        : Dong sub Kim
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit         '��� ���ؿ��� �ش��ϴ� ����� ��� ������ ���������� �����ϵ��� �� �� ����մϴ�.

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_MM060H       As SAPbouiCOM.DBDataSource    '������ 'Represents a SAP Business One database table to be attached to a form.
Private oDS_PS_MM060L       As SAPbouiCOM.DBDataSource    '��϶���
Private oDS_PS_MM060M       As SAPbouiCOM.DBDataSource

Private oLast_Item_UID   As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID    As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row    As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��
Private oSeq&
Private TmpCode As String

'*******************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM060.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '��Ʈ������ Ÿ��Ʋ���̿� �����̸� ����
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_MM060_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    ' ������ �Ӽ� � ��ü ������ �����մϴ�

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////D////////////////////////////////////////////////
        
    oForm01.Freeze True
    Call CreateItems
    
    oForm01.EnableMenu ("1283"), True      '// ����
    oForm01.EnableMenu ("1293"), True      '// �����
    oForm01.EnableMenu ("1287"), True      '// ����
    oForm01.EnableMenu ("1284"), False     '// ���
    
    Call ComboBox_Setting
    Call CF_ChooseFromList
    Call Initial_Setting
    Call FormItemEnabled
    Call FormClear '//UDO����϶�
    Call AddMatrixRow(0, oMat01.RowCount, True) '//UDO����϶�
    oForm01.Update      ' Reloads the form from the application.
                        ' Use this method to update an object from the database.

    oForm01.Freeze False
       
    oForm01.Visible = True  ' Indicates whether the form is visible.
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
    Dim sQry As String
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        '// Input Code, Name
                        oDS_PS_MM060H.setValue "Code", 0, Trim(oDS_PS_MM060H.GetValue("U_ItmMsort", 0))
                        oDS_PS_MM060H.setValue "Name", 0, Trim(oDS_PS_MM060H.GetValue("U_ItmMname", 0))
                        TmpCode = Trim(oDS_PS_MM060H.GetValue("U_ItmMsort", 0))
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                ' ��з�
                If oForm01.Items("U_ItmBsort").Specific.Value = "" Then
                    If pval.ItemUID = "U_ItmBsort" And pval.CharPressed = 9 Then
                        oForm01.Items("U_ItmBsort").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                '�Һз�
                If oForm01.Items("U_ItmMsort").Specific.Value = "" Then
                    If pval.ItemUID = "U_ItmMsort" And pval.CharPressed = 9 Then
                        If oForm01.Items("U_ItmBsort").Specific.Value = "" Then
                           ErrNum = 1
                           Call Item_Error_Message(ErrNum)
                        Else
                           oForm01.Items("U_ItmMsort").Click ct_Regular
                           Sbo_Application.ActivateMenuItem ("7425")
                           BubbleEvent = False
                        End If
                    End If
                End If
                
                '�Է���
                If oForm01.Items("CntcCode").Specific.Value = "" Then
                    If pval.ItemUID = "CntcCode" And pval.CharPressed = 9 Then
                        oForm01.Items("CntcCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
                 oLast_Item_UID = pval.ItemUID

            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
            '//��Ʈ������ ������ �ε�
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
              If pval.ItemUID = "1" Then
              
            ' HISTORY table ī��
                If (oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Or oForm01.Mode = fm_OK_MODE) And pval.Action_Success = True Then
                    Call PS_MM060_MTX02
                Else
                    FormItemEnabled
                    AddMatrixRow 1, oMat01.RowCount, True
                End If
                
                If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                    oForm01.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = False Then
                    FormItemEnabled
                    AddMatrixRow 1, oMat01.RowCount, True
                End If
              End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                ' ��з�
                If pval.ItemUID = "U_ItmBsort" And pval.ItemChanged = True Then
                    sQry = "Select Name From [@PSH_ItmBsort] Where Code = '" & Trim(oForm01.Items("U_ItmBsort").Specific.Value) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("ItmBname").Specific.Value = Trim(oRecordSet01.Fields(0).Value)
                End If
                
                ' �ߺз�
                If pval.ItemUID = "U_ItmMsort" And pval.ItemChanged = True Then
                    Call PS_MM060_MTX01
                End If
                
                ' �Է���
                If pval.ItemUID = "CntcCode" And pval.ItemChanged = True Then
                    sQry = "Select U_FULLNAME, U_MSTCOD From [OHEM] Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.Value) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("CntcName").Specific.Value = Trim(oRecordSet01.Fields(0).Value)
                End If
                
            Case et_MATRIX_LOAD: '//11
                AddMatrixRow 1, oMat01.VisualRowCount, True
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
                oLast_Item_UID = pval.ItemUID
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    If (pval.BeforeAction = True) Then      '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
            Case "1293"     '�����
        End Select
        
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1281": 'ã��
                FormItemEnabled
            Case "1282": '�߰�
                FormItemEnabled
                FormClear
                AddMatrixRow 0, oMat01.RowCount, True
                oForm01.Items("U_ItmBsort").Click ct_Collapsed
                
            Case "1287": '����
                oForm01.Freeze True
                oDS_PS_MM060H.setValue "Code", 0, ""
                oDS_PS_MM060H.setValue "Name", 0, ""
                oDS_PS_MM060H.setValue "U_Major", 0, ""
                oDS_PS_MM060H.setValue "U_MajorNm", 0, ""
                oDS_PS_MM060H.setValue "U_Remark", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_MM060L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
                 FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("Spec1").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
                        AddMatrixRow 1, oMat01.RowCount, True
                    End If
                End If
                
            Case "1293"     '�����
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.Value = i
                    Next i
                    oMat01.FlushToDataSource    ' DBDataSource�� ���ڵ尡 ���� �� �����.
                    Call oDS_PS_MM060L.RemoveRecord(oDS_PS_MM060L.Size - 1)     ' ���ڵ� �� ���� �����.
                    oMat01.LoadFromDataSource   ' DBDataSource�� ��Ʈ������ �ø���
                    If oMat01.RowCount = 0 Then
                        Call AddMatrixRow(1, 0, True)
                    Else
                        If Trim(oDS_PS_MM060L.GetValue("U_Spec1", oMat01.RowCount - 1)) <> "" Then
                            Call AddMatrixRow(1, oMat01.RowCount, True)

                        End If
                    End If
                End If
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
                FormItemEnabled
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
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM060H = oForm01.DataSources.DBDataSources("@PS_MM060H")
    Set oDS_PS_MM060L = oForm01.DataSources.DBDataSources("@PS_MM060L")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Set oRecordSet01 = Nothing
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//�޺��� �⺻������
    Exit Sub
ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList ����
    Exit Sub
CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub Initial_Setting()
On Error GoTo Initial_Setting_Error
    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD()
Exit Sub
Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "Initial_Setting_Error: " & Err.Number & " _ " & Err.Description, bmt_Short, True
End Sub


Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//����忡���� �����ۼ���
        oForm01.Items("U_ItmBsort").Enabled = True
        oForm01.Items("U_ItmMsort").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oMat01.Columns("Price1").Editable = True
        oMat01.Columns("Price2").Editable = True
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//����忡���� �����ۼ���
        oForm01.Items("U_ItmBsort").Enabled = True
        oForm01.Items("U_ItmMsort").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oMat01.Columns("Price1").Editable = False
        oMat01.Columns("Price2").Editable = False
        
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//����忡���� �����ۼ���
        oForm01.Items("U_ItmBsort").Enabled = False
        oForm01.Items("U_ItmMsort").Enabled = False
        oForm01.Items("CntcCode").Enabled = False
        oMat01.Columns("Price1").Editable = True
        oMat01.Columns("Price2").Editable = True
    End If
    Exit Sub
FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error
   Select Case oSeq
       Case 0:
            oMat01.AddRow
            oDS_PS_MM060L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
       Case 1:
            oDS_PS_MM060L.InsertRecord oRow
            oDS_PS_MM060L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
   End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM060'", "")
    If DocNum = 0 Then
'        oForm01.Items("DocEntry").Specific.String = 1
        oDS_PS_MM060H.setValue "DocEntry", 0, "1"
    Else
'        oForm01.Items("DocEntry").Specific.String = DocNum
        oDS_PS_MM060H.setValue "DocEntry", 0, DocNum    ' ȭ�鿡 ������ �ȵǱ� ����
    End If
    Exit Sub
FormClear_Error:
    Sbo_Application.SetStatusBarMessage "FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function DataValidCheck() As Boolean
On Error GoTo DataValidCheck_Error
    '//��ȿ���˻�
    Exit Function
DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub MTX01()
On Error GoTo MTX01_Error
    '//��Ʈ������ ������ �ε�
    Exit Sub
MTX01_Error:
Sbo_Application.SetStatusBarMessage "MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    oForm01.Freeze True
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_MM060H.GetValue("U_ItmBsort", 0) = "" Or oDS_PS_MM060H.GetValue("U_ItmMsort", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select
    oForm01.Freeze False
    HeaderSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��з� �Ǵ� �ߺз��� �ʼ��Է� �����Դϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    Dim sQry          As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    ErrNum = 0
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// ȭ����� ��Ʈ������ �Էµ� ������ ��� ������ͼҽ��� �ѱ��
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.FlushToDataSource
    
    '// ����
    If oMat01.VisualRowCount <= 1 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '// �Ǹ������� �����͸� �����ϴ� ������ ���� �߰� �Ұ�쿡 ������ͼҽ���
    '// �̹� �����Ͱ� �� �ֱ� ������ ����ÿ��� ������ ��(DB������ �ҽ���)�� �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    If oMat01.VisualRowCount > 0 Then
        If oDS_PS_MM060L.GetValue("U_Spec1", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_MM060L.RemoveRecord oMat01.VisualRowCount - 1
        End If
    End If
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���� �����Ͽ����� DB������ �ҽ��� �ٽ� �����´�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    oMat01.LoadFromDataSource
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "���� �����Ͱ� �����ϴ�. Ȯ���ϼ���.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    Dim i$
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "Mat01"
            If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 2) And _
            Trim(oMat01.Columns("Minor").Cells(oRow).Specific.Value) <> "" Then
                oMat01.FlushToDataSource
                AddMatrixRow 1, oMat01.RowCount, True
                oMat01.Columns("Spec1").Cells(oRow).Click ct_Regular
            End If
    End Select
    Set oRecordSet01 = Nothing
End Sub
Private Function Item_Error_Message(ErrNum As Integer)
On Error GoTo Item_Error_Message

    If ErrNum = 1 Then
        GoTo Item_Error_Message
    End If

Item_Error_Message:

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "��з��� ���� �Է��ϼ���!", "E"
    End If
    
End Function

Private Sub PS_MM060_MTX01()
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim oRecordSet02 As SAPbobsCOM.Recordset
    Dim i As Integer
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Freeze True     ' Freezes and unfreezes the form. When frozen, no changes in the form are displayed.
                            ' Use this method to update the form without causing flickering.

    sQry = "Select U_CodeName From [@PSH_ITMMSORT] Where U_rCode = '" & Trim(oForm01.Items("U_ItmBsort").Specific.Value) & "'"
    sQry = sQry + " And U_Code = '" & Trim(oForm01.Items("U_ItmMsort").Specific.Value) & "'"
    oRecordSet01.DoQuery sQry
    oForm01.Items("ItmMname").Specific.Value = Trim(oRecordSet01.Fields(0).Value)
    
    sQry = "Select U_Minor From [@PS_SY001L] Where Code = 'M005'"
    oRecordSet02.DoQuery sQry
    
    oDS_PS_MM060L.Clear
    oMat01.Clear
    oMat01.FlushToDataSource
    
    i = 0
    Do Until oRecordSet02.EOF
        oDS_PS_MM060L.InsertRecord i
        oDS_PS_MM060L.setValue "U_LineNum", i, i + 1
        oDS_PS_MM060L.setValue "U_Spec1", i, Trim(oRecordSet02.Fields(0).Value)
        oDS_PS_MM060L.setValue "Code", i, Trim(oDS_PS_MM060H.GetValue("U_ItmMsort", 0))
        i = i + 1
        oRecordSet02.MoveNext
    Loop
    
    oMat01.LoadFromDataSource
    oForm01.Freeze False
End Sub

Private Sub PS_MM060_MTX02()
  ' ���� ���̺� ��/�ߺз� �ڵ� UPDATE
    Dim Param01 As String
    Dim Param02 As String
    Dim Query01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    oForm01.Freeze True
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Param01 = Format(Now, "yyyymmdd")
    Param02 = Trim(TmpCode)
    Query01 = "EXEC PS_MM060_02 '" & Param01 & "', '" & Param02 & "'"
    
    Call oRecordSet01.DoQuery(Query01)
  
  ' �����丮 ���̺� INSERT
    Dim Param11 As String
    Dim Query02 As String
    Dim oRecordSet02 As SAPbobsCOM.Recordset
    
    Set oRecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Param11 = TmpCode
    Query02 = "EXEC PS_MM060_01 '" & Param11 & "'"
    Call oRecordSet02.DoQuery(Query02)
    oForm01.Freeze False
End Sub