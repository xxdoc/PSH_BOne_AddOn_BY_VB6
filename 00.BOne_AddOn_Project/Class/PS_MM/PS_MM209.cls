VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM209"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM209.cls
'//  Module         : MM
'//  Description    : ���Ҹ�����
'//  FormType       : PS_MM209
'//  Create Date    : 2011.02.17
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_MM209H       As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_MM209L       As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM209.srf")
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
    
    oFormUniqueID01 = "PS_MM209_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    'ȭ��Ű��(ȭ�鿡�� ����Ű���� ��� �ִ� �������� Uid��)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False        '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
    oForm01.EnableMenu ("1284"), False        '// ���
    oForm01.EnableMenu ("1293"), False        '// �����
        
    oForm01.Update
    oForm01.Freeze False
    oForm01.Visible = True
    
    Set oXmlDoc01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    If (oForm01 Is Nothing) = False Then
        Set oForm01 = Nothing
    End If
    MDC_Com.MDC_GF_Message "LoadForm_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
'    Set oDS_PS_MM209H = oForm01.DataSources.DBDataSources("@PS_MM209H")
'    Set oDS_PS_MM209L = oForm01.DataSources.DBDataSources("@PS_MM209L")
    
    '// ��Ʈ���� ��ü �Ҵ�
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("StrDate", dt_DATE, 10)
    oForm01.Items("StrDate").Specific.DataBind.SetBound True, "", "StrDate"
    oForm01.DataSources.UserDataSources.Item("StrDate").VALUE = Format(Now, "YYYYMMDD")

    Call oForm01.DataSources.UserDataSources.Add("EndDate", dt_DATE, 10)
    oForm01.Items("EndDate").Specific.DataBind.SetBound True, "", "EndDate"
    oForm01.DataSources.UserDataSources.Item("EndDate").VALUE = Format(Now, "YYYYMMDD")

    '//üũ�ڽ� ó��
    Dim oChkBox As SAPbouiCOM.CheckBox

    Call oForm01.DataSources.UserDataSources.Add("Check01", dt_SHORT_TEXT, 1)
    
    Set oChkBox = oForm01.Items("ChkBox01").Specific
    oChkBox.ValOn = "Y"
    oChkBox.ValOff = "N"
    oChkBox.DataBind.SetBound True, , "Check01"

    oForm01.DataSources.UserDataSources("Check01").VALUE = "N" '��üũ�� ���� �ְ� ���� �ε�
    
    Call oForm01.DataSources.UserDataSources.Add("Check02", dt_SHORT_TEXT, 1)
    
    Set oChkBox = oForm01.Items("ChkBox02").Specific
    oChkBox.ValOn = "Y"
    oChkBox.ValOff = "N"
    oChkBox.DataBind.SetBound True, , "Check02"

    oForm01.DataSources.UserDataSources("Check02").VALUE = "N" '��üũ�� ���� �ְ� ���� �ε�

    Set oChkBox = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//�޺��� �⺻������
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// �������
    Set oCombo = oForm01.Items("AcctCode").Specific
    sQry = "SELECT BPLId, BPLName From OBPL"
    oRecordSet01.DoQuery sQry
    Call oCombo.ValidValues.Add("11506100", "�����")
    Call oCombo.ValidValues.Add("11502100", "��ǰ")
    Call oCombo.ValidValues.Add("11501100", "��ǰ")
    Call oCombo.ValidValues.Add("11507100", "����ǰ")
    Call oCombo.ValidValues.Add("11503100", "���ǰ")
    Call oCombo.Select(0, psk_Index)
    
    '// â��
    Set oCombo = oForm01.Items("WhsCode").Specific
    sQry = "SELECT WhsCode, WhsName From OWHS"
    oRecordSet01.DoQuery sQry
    
    Call oCombo.ValidValues.Add("000", "��ü")
    
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oCombo.Select(0, psk_Index)

    
    '// �����
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From OBPL"
    oRecordSet01.DoQuery sQry
    Dim i As Integer
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
        i = i + 1
    Loop
    
    Call oCombo.ValidValues.Add(6, "�Ȱ�+�������")
    
    Call oCombo.Select(0, psk_Index)
    


    '// ������ ��з�
'    Set oCombo = oForm01.Items("ItmGrp").Specific
'    sQry = "SELECT ItmsGrpCod, ItmsGrpNam From [OITB] where ItmsGrpCod<>'103'"
'    oRecordSet01.DoQuery sQry
'    Call oCombo.ValidValues.Add("001", "��ü")
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
'    Call oCombo.Select(0, psk_Index)
        
    '// ������ ��з�
    Set oCombo = oForm01.Items("ItmBsort").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code"
    oRecordSet01.DoQuery sQry
    Call oCombo.ValidValues.Add("001", "��ü")
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oCombo.Select(0, psk_Index)
    
    
    '// ������ �ߺз�
    Set oCombo = oForm01.Items("ItmMsort").Specific
    sQry = "SELECT U_Code,U_CodeName FROM [@PSH_ITMMSORT] Order by U_Code"
    oRecordSet01.DoQuery sQry
    
    If oForm01.Items("ItmMsort").Specific.ValidValues.Count = 0 Then
        Call oCombo.ValidValues.Add("00001", "��ü")
    End If
    
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oCombo.Select(0, psk_Index)
    
    
    Set oCombo = oForm01.Items("Gubun").Specific
    Call oCombo.ValidValues.Add("1", "����")
    Call oCombo.ValidValues.Add("2", "����")
    Call oCombo.Select(0, psk_Index)
    
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Dim sCount&
    Dim sSeq&
    Dim ItmBsort As String
    
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
'                        If HeaderSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                    End If
                
                ElseIf pval.ItemUID = "Btn01" Then '��¹�ư Ŭ����
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call Print_Query
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//���
'                    If pval.ItemUID = "ItmBsort" Then
'                        If oForm01.Items("ItmBsort").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
'                    If pval.ItemUID = "SItemCode" Then
'                        If oForm01.Items("SItemCode").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
'                    If pval.ItemUID = "EItemCode" Then
'                        If oForm01.Items("EItemCode").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
                    '//����
'                    If pval.ItemUID = "Mat01" Then
'                        If pval.ColUID = "PP070No" Then
'                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.Value = "" Then
'                                Sbo_Application.ActivateMenuItem ("7425")
'                                BubbleEvent = False
'                            End If
'                        End If
'                    End If
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
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
'                If pval.ItemUID = "1" Then
'                    If oForm01.Mode = fm_ADD_MODE Then
'                        oForm01.Mode = fm_OK_MODE
'                        Call Sbo_Application.ActivateMenuItem("1282")
'                    ElseIf oForm01.Mode = fm_OK_MODE Then
'                        FormItemEnabled
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
'                    End If
'                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                
                If pval.ItemUID = "ItmBsort" Then
                    Call oForm01.Freeze(True)
                    sCount = oForm01.Items("ItmMsort").Specific.ValidValues.Count
                    sSeq = sCount
                    For i = 1 To sCount
                        oForm01.Items("ItmMsort").Specific.ValidValues.Remove sSeq - 1, psk_Index
                        sSeq = sSeq - 1
                    Next i

                    ItmBsort = Trim(oForm01.Items("ItmBsort").Specific.VALUE)
                    Set oCombo = oForm01.Items("ItmMsort").Specific
                    sQry = "SELECT U_Code,U_CodeName FROM [@PSH_ITMMSORT] Where U_rCode = '" & ItmBsort & "' Order by U_Code"
                    oRecordSet01.DoQuery sQry
                    
                    If oForm01.Items("ItmMsort").Specific.ValidValues.Count = 0 Then
                        Call oCombo.ValidValues.Add("00001", "��ü")
                    End If
                    
                    Do Until oRecordSet01.EOF
                        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                        oRecordSet01.MoveNext
                    Loop
                    Call oCombo.Select(0, psk_Index)

                    Set oCombo = Nothing
                    Set oRecordSet01 = Nothing
                    
                    Call oForm01.Freeze(False)
                End If
                
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
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
'                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    If ErrNum = 101 Then
        ErrNum = 0
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
        BubbleEvent = False
    Else
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1285": '����
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
            Case "1285": '����
            Case "1293": '�����
            Case "1281": 'ã��
            Case "1282": '�߰�
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

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
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
    
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
        
        End Select
    End If
     
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("StrDate").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("EndDate").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "�������� ������ �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "�������� ����� �ʼ������Դϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub Print_Query()
On Error GoTo Print_Query_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry        As String
    
    Dim ItmGrp          As String
    Dim ItmBsort        As String
    Dim ItmMsort        As String
    Dim BPLID           As String
    Dim StrDate         As String
    Dim EndDate         As String
    Dim SItemCode       As String
    Dim EITemCode       As String
    Dim AcctCode        As String
    Dim WhsCode         As String
    Dim ChkBox          As String, ChkBox02$
    Dim Gubun           As String
    
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ ��...", 100, False)
            
    Call ConnectODBC
    
    '// ��ȸ���ǹ�
'    ItmGrp = Trim(oForm01.Items("ItmGrp").Specific.Selected.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBsort").Specific.VALUE)
    ItmMsort = Trim(oForm01.Items("ItmMsort").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    StrDate = Trim(oForm01.Items("StrDate").Specific.VALUE)
    EndDate = Trim(oForm01.Items("EndDate").Specific.VALUE)
    AcctCode = Trim(oForm01.Items("AcctCode").Specific.Selected.VALUE)
    WhsCode = Trim(oForm01.Items("WhsCode").Specific.Selected.VALUE)
    ChkBox = Trim(oForm01.DataSources.UserDataSources("Check01").VALUE)
    ChkBox02 = Trim(oForm01.DataSources.UserDataSources("Check02").VALUE)
    Gubun = Trim(oForm01.Items("Gubun").Specific.Selected.VALUE)
'    SItemCode = Trim(oForm01.Items("SItemCode").Specific.VALUE)
'    EITemCode = Trim(oForm01.Items("EItemCode").Specific.VALUE)
    
    If StrDate = "" Then StrDate = "19000101"
    If EndDate = "" Then EndDate = "21001231"
'    If SItemCode = "" Then SItemCode = "%"
'    If ItmGrp = "001" Then ItmGrp = "%"
    If ItmBsort = "001" Then ItmBsort = "%"
    If ItmMsort = "00001" Then ItmMsort = "%"
    
'    If BPLId = "0" Then
'        BPLId = "%"
'    ElseIf BPLId = "1" Then
'        BPLId = "%1"
'    ElseIf BPLId = "2" Then
'        BPLId = "%2"
'    ElseIf BPLId = "3" Then
'        BPLId = "%3"
'    ElseIf BPLId = "4" Then
'        BPLId = "%4"
'    ElseIf BPLId = "5" Then
'        BPLId = "%5"
'    End If
        
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    
    If Gubun = "1" Then
        WinTitle = "[PS_MM209] ���Ҹ�����"
        ReportName = "PS_MM209_01.RPT"
    ElseIf Gubun = "2" Then
        WinTitle = "[PS_MM209] ���Ҹ�����(����)"
        ReportName = "PS_MM209_04.RPT"
    End If
    
    '// Formula �����ʵ�
    ReDim gRpt_Formula(5)
    ReDim gRpt_Formula_Value(5)
    
    gRpt_Formula(1) = "StrDate": gRpt_Formula_Value(1) = IIf(StrDate = "", "All", Format(StrDate, "0000-00-00"))
    
    gRpt_Formula(2) = "EndDate": gRpt_Formula_Value(2) = IIf(EndDate = "", "All", Format(EndDate, "0000-00-00"))
    
    gRpt_Formula(3) = "BPLId"
    If BPLID = "6" Then
        gRpt_Formula_Value(3) = "�Ȱ�+��� �����"
    Else
        sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" & BPLID & "'"
        Call oRecordSet.DoQuery(sQry)
        gRpt_Formula_Value(3) = oRecordSet.Fields(0).VALUE
    End If
    
    
    gRpt_Formula(4) = "AcctCode"
    gRpt_Formula_Value(4) = AcctCode
    
    gRpt_Formula(5) = "ChkBox"
    gRpt_Formula_Value(5) = ChkBox
    
    sQry = "SELECT WhsName From OWHS where WhsCode = '" & WhsCode & "'"
    oRecordSet.DoQuery sQry
    gRpt_Formula(5) = "WhsName"
    gRpt_Formula_Value(5) = IIf(WhsCode = "000", "��ü", oRecordSet.Fields(0).VALUE)
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    If Gubun = "1" Then
        '//���� ����
        '/ Procedure ����"
        If ChkBox02 = "Y" Then
            'sQry = "EXEC [PS_MM209_03] '" & BPLId & "','" & StrDate & "','" & EndDate & "','" & AcctCode & "', '" & WhsCode & "', '" & ChkBox & "', '" & ItmBsort & "', '" & ItmMsort & "'"
            '����������
            sQry = "EXEC [PS_MM209_10] '" & BPLID & "','" & StrDate & "','" & EndDate & "','" & AcctCode & "', '" & WhsCode & "', '" & ChkBox & "', '" & ItmBsort & "', '" & ItmMsort & "'"
        Else
            sQry = "EXEC [PS_MM209_02] '" & BPLID & "','" & StrDate & "','" & EndDate & "','" & AcctCode & "', '" & WhsCode & "', '" & ChkBox & "', '" & ItmBsort & "', '" & ItmMsort & "'"
            
        End If
    Else
        '//��������
        sQry = "EXEC [PS_MM209_04] '" & BPLID & "','" & StrDate & "','" & EndDate & "','" & AcctCode & "', '" & WhsCode & "', '" & ChkBox & "', '" & ItmBsort & "', '" & ItmMsort & "'"
    End If
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query�� �������� 'Y'��...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        GoTo Print_Query_Error
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
'    If ErrNum = 1 Then
'        MDC_Com.MDC_GF_Message "����� �����Ͱ� �����ϴ�. Ȯ���� �ּ���.", "E"
'    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
End Sub

