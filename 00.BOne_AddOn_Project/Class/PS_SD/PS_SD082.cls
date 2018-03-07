VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD082"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_SD082.cls
'//  Module         : SD
'//  Description    : �����ѵ� �ʰ�����
'//  FormType       : PS_SD082
'//  Create Date    : 2010.10.22
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Public oMat02              As SAPbouiCOM.Matrix
Private oDS_PS_SD082L As SAPbouiCOM.DBDataSource    '��϶���
Private oDS_PS_SD082M As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��

Private oLast_Mode&
Private oSeq&

'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD082.srf")
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
    
    oFormUniqueID01 = "PS_SD082_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//���߰�
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '�� �Ҵ�
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
        
    oForm01.Freeze True
        
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    oForm01.Mode = fm_OK_MODE
    Call LoadCaption
    
    oForm01.EnableMenu ("1283"), False        '// ����
    oForm01.EnableMenu ("1286"), False         '// �ݱ�
    oForm01.EnableMenu ("1287"), False        '// ����
    oForm01.EnableMenu ("1284"), False         '// ���
    oForm01.EnableMenu ("1293"), False         '// �����
        
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

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim TempForm01 As Variant
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
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
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If oForm01.Mode = fm_UPDATE_MODE Then
                        Call Update_SD080(pval)
                        oForm01.Mode = fm_OK_MODE
                        oMat01.Clear
                        oDS_PS_SD082L.Clear
                        oMat02.Clear
                        oDS_PS_SD082M.Clear
                        Call LoadCaption
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                        oForm01.Close
                    End If
                ElseIf pval.ItemUID = "Btn02" Then
                    Call LoadData
                    oForm01.Mode = fm_OK_MODE
                    Call LoadCaption
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
'et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_CLICK: '//6
                If pval.ItemUID = "Radio01" Then
                    oForm01.Freeze True
                    oForm01.Settings.MatrixUID = "Mat01"
                    oForm01.Settings.EnableRowFormat = True
                    oForm01.Settings.Enabled = True
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "Radio02" Then
                    oForm01.Freeze True
                    oForm01.Settings.MatrixUID = "Mat02"
                    oForm01.Settings.EnableRowFormat = True
                    oForm01.Settings.Enabled = True
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "Mat01" Then
                    If pval.ColUID = "LineNum" Then
                        LoadData_Mat02 (Trim(oDS_PS_SD082L.GetValue("U_ColReg02", pval.Row - 1)))
                    ElseIf pval.ColUID = "Check" Then
                        oForm01.Freeze True
                        oMat01.FlushToDataSource
                        For i = 0 To oMat01.VisualRowCount - 1
                            If Trim(oDS_PS_SD082L.GetValue("U_ColReg01", i)) = "Y" Then
                                oForm01.Mode = fm_UPDATE_MODE
                                Call LoadCaption
                                oForm01.Freeze False
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Next i
                        oForm01.Mode = fm_OK_MODE
                        Call LoadCaption
                        oForm01.Freeze False
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
                oForm01.Freeze True
                
                oForm01.Items("Mat01").Top = 50
                oForm01.Items("Mat01").Left = 6
                oForm01.Items("Mat01").Width = oForm01.Width * 0.4 - 6
                oForm01.Items("Mat01").Height = oForm01.Height - 110

                oForm01.Items("Mat02").Top = oForm01.Items("Mat01").Top
                oForm01.Items("Mat02").Left = oForm01.Width * 0.4 + 6 + 10
                oForm01.Items("Mat02").Width = oForm01.Width * 0.6 - 6 - 22
                oForm01.Items("Mat02").Height = oForm01.Height - 110
                
                oForm01.Items("Radio01").Left = 6
                oForm01.Items("Radio02").Left = oForm01.Width * 0.4 + 6 + 10
                
                oMat01.AutoResizeColumns
                oMat02.AutoResizeColumns

'                oMat01.Columns("Check").Width = 40
'                oMat01.Columns("DocNum").Width = 60
'                oMat01.Columns("BPLId").Width = 50
'                oMat01.Columns("CntcCode").Width = 60
'                oMat01.Columns("DocDate").Width = 80
'
'                oMat02.Columns("CardCode").Width = 80
'                oMat02.Columns("CardName").Width = 80
'                oMat02.Columns("RequestP").Width = 80
'                oMat02.Columns("CreditP").Width = 80
'                oMat02.Columns("MiSuP").Width = 80
'                oMat02.Columns("Balance").Width = 80
'                oMat02.Columns("OutPreP").Width = 80
'                oMat02.Columns("Comment").Width = 80
                
                oForm01.Freeze False
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oMat02 = Nothing
                Set oDS_PS_SD082L = Nothing
                Set oDS_PS_SD082M = Nothing
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
            Case "1288", "1289", "1290", "1291": '���ڵ��̵���ư
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '���
            Case "1286": '�ݱ�
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//�۾�
    ElseIf (eventInfo.BeforeAction = False) Then
        '//�۾�
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//������� �ҽ� ��ü �Ҵ�
    Set oDS_PS_SD082L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_SD082M = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    
    '// ��Ʈ���� ��ü �Ҵ�
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    
    oForm01.DataSources.UserDataSources.Add "Radio01", dt_SHORT_TEXT, 1
    oForm01.Items("Radio01").Specific.DataBind.SetBound True, "", "Radio01"
    
    oForm01.DataSources.UserDataSources.Add "Radio02", dt_SHORT_TEXT, 1
    oForm01.Items("Radio02").Specific.DataBind.SetBound True, "", "Radio02"
    
    oForm01.Items("Radio01").Specific.GroupWith "Radio02"
    
'    oDS_PS_SD082L.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
'    oDS_PS_SD082L.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
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
    
    Call oForm01.DataSources.UserDataSources.Add("OkYN", dt_SHORT_TEXT, 1)
    oForm01.Items("OkYN").Specific.DataBind.SetBound True, "", "OkYN"
    
    '// ���λ���
    oForm01.Items("OkYN").Specific.ValidValues.Add "Y", "����"
    oForm01.Items("OkYN").Specific.ValidValues.Add "N", "�̽���"
    oForm01.Items("OkYN").Specific.Select "1", psk_Index
    
    '// �����
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oRecordSet01.MoveNext
    Loop
    
    '// �����
    sQry = "Select empID, lastName + firstName From OHEM Order by empID"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("CntcCode").ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
        oRecordSet01.MoveNext
    Loop
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//���̵� ����� ����
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//���̵� ��� ����
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
    '//���̵� �μ� ����
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub LoadCaption()
On Error GoTo LoadCaption_Error
    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "Ȯ��"
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "Ȯ��"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("Btn01").Specific.Caption = "����"
    End If

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, OkYN$, DocNum$
        
    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
    OkYN = Trim(oForm01.DataSources.UserDataSources("OkYN").Value)
    
    If OkYN = "" Then OkYN = "%"
        
    sQry = "EXEC [PS_SD082_01] '" & BPLId & "','" & OkYN & "','" & DocNum & "','01'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_SD082L.Clear
    
    oMat02.Clear
    oDS_PS_SD082M.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_SD082L.Size Then
            oDS_PS_SD082L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_SD082L.Offset = i
        oDS_PS_SD082L.setValue "U_LineNum", i, i + 1
        oDS_PS_SD082L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("DocNum").Value)
        oDS_PS_SD082L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_BPLId").Value)
        oDS_PS_SD082L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("U_CntcCode").Value)
        oDS_PS_SD082L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("U_DocDate").Value)

        oRecordSet01.MoveNext
        ProgBar01.Value = ProgBar01.Value + 1
        ProgBar01.Text = ProgBar01.Value & "/" & oRecordSet01.RecordCount & "�� ��ȸ��...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData_Mat02(ByVal sDocNum As String)
On Error GoTo LoadData_Mat02_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Dim BPLId$, OkYN$, DocNum$
        
    BPLId = Trim(oForm01.Items("BPLId").Specific.Value)
    OkYN = Trim(oForm01.DataSources.UserDataSources("OkYN").Value)
    
    sQry = "EXEC [PS_SD082_01] '" & BPLId & "','" & OkYN & "','" & sDocNum & "','02'"
    oRecordSet01.DoQuery sQry
    
    oMat02.Clear
    oDS_PS_SD082M.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "��ȸ ����� �����ϴ�. Ȯ���ϼ���.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("��ȸ����!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_SD082M.Size Then
            oDS_PS_SD082M.InsertRecord (i)
        End If
        
        oMat02.AddRow
        oDS_PS_SD082M.Offset = i
        oDS_PS_SD082M.setValue "U_LineNum", i, i + 1
        oDS_PS_SD082M.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("U_CardCode").Value)
        oDS_PS_SD082M.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("U_CardName").Value)
        oDS_PS_SD082M.setValue "U_ColSum01", i, Trim(oRecordSet01.Fields("U_RequestP").Value)
        oDS_PS_SD082M.setValue "U_ColSum02", i, Trim(oRecordSet01.Fields("U_CreditP").Value)
        oDS_PS_SD082M.setValue "U_ColSum03", i, Trim(oRecordSet01.Fields("U_MiSuP").Value)
        oDS_PS_SD082M.setValue "U_ColSum04", i, Trim(oRecordSet01.Fields("U_Balance").Value)
        oDS_PS_SD082M.setValue "U_ColSum05", i, Trim(oRecordSet01.Fields("U_OutPreP").Value)
        oDS_PS_SD082M.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_Comment").Value)

        oRecordSet01.MoveNext
        ProgBar01.Value = ProgBar01.Value + 1
        ProgBar01.Text = ProgBar01.Value & "/" & oRecordSet01.RecordCount & "�� ��ȸ��...!"
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Mat02_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Mat02_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Function Update_SD080(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Update_JakNum_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    Dim DocNum$, OkDate$
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.RowCount - 1
        If Trim(oDS_PS_SD082L.GetValue("U_ColReg01", i)) = "Y" Then
            DocNum = Trim(oDS_PS_SD082L.GetValue("U_ColReg02", i))
            OkDate = Format(Now, "YYYYMMDD")
            
            sQry = "UPDATE [@PS_SD080H] "
            sQry = sQry & "SET "
            sQry = sQry & "U_OkYN = 'Y', "
            sQry = sQry & "U_OkDate = '" & OkDate & "'"
            sQry = sQry & "Where DocNum = '" & DocNum & "'"
        
            RecordSet01.DoQuery sQry
        End If
    Next

    MDC_Com.MDC_GF_Message "�����ѵ� �ʰ����� �Ϸ�!", "S"
    
    Set RecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_JakNum_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Update_JakNum_Error:" & Err.Number & " - " & Err.Description, "E"
End Function