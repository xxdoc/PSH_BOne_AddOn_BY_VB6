VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_FTP.cls
'//  Module         : FI
'//  Description    : ���ձ��� FTP ��������
'//  FormType       : PS_FTP
'//  Create Date    : 2013.07.19
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_FTPH       As SAPbouiCOM.DBDataSource    '������
Private oDS_PS_FTPL       As SAPbouiCOM.DBDataSource    '��϶���

Private oLast_Item_UID      As String                     'Ŭ�������� ������ ������ ������ Uid��
Private oLast_Col_UID       As String                     '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Col�� Uid��
Private oLast_Col_Row       As Long                       '�������������� ��Ʈ�����ϰ�쿡 ������ ���õ� Row��


Dim hOpen As Long, hConnection As Long
Dim dwType As Long
Const ASCII As Long = FTP_TRANSFER_TYPE_ASCII
Const BINARY As Long = FTP_TRANSFER_TYPE_BINARY
Const FTP_SERVER As String = "192.1.11.3"  '"192.1.11.3" '7���� �׽�Ʈ��
Const FTP_USER As String = "ftpadm"
Const FTP_PASSWORD As String = "psc1004"

Private oLast_Mode&



'****************************************************************************************************************
' .srf ���Ϸκ��� ���� �ε��Ѵ�.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByVal oReqNo01 As String, Optional ByVal oBANFN01 As String, Optional ByVal oBNFPO01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FTP.srf")
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
    
    oFormUniqueID01 = "PS_FTP_" & GetTotalFormsCount
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
    
    oForm01.Items("ReqNo").Specific.VALUE = oReqNo01
    oForm01.Items("E_BANFN").Specific.VALUE = oBANFN01
    oForm01.Items("E_BNFPO").Specific.VALUE = oBNFPO01
    
    
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
'    Set oDS_PS_FTPH = oForm01.DataSources.DBDataSources("@PS_FTPH")
'    Set oDS_PS_FTPL = oForm01.DataSources.DBDataSources("@PS_FTPL")
    
    
    
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
                
                ElseIf pval.ItemUID = "Btn01" Then '���� Ŭ����
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call FTP_Trans
                    End If
                
                ElseIf pval.ItemUID = "Btn02" Then '����ã��
                    Call FileSearch
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
                    If pval.ItemUID = "SItemCode" Then
                        If oForm01.Items("SItemCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "EItemCode" Then
                        If oForm01.Items("EItemCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
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
    
    Select Case True
        Case Trim(oForm01.Items("ReqNo").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("E_BANFN").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("FileName").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "û����ȣ(B1)�� �����ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "���ſ�û��ȣ(R3)�� �����ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "������ ���ϸ��� �����ϴ�. Ȯ���Ͽ� �ֽʽÿ�.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub FTP_Trans()
On Error GoTo FTP_Trans_Error
   '
'����
'UpLoad Local_FileURL , Remote_FileURL , ASCII or BINARY
'DownLoad Remote_FileURL , Local_FileURL , ASCII or BINARY
'

Dim Dir As String
Dim oFileName  As String
Dim FileName As String
Dim pos As Integer
Dim ReturnValue As String
Dim BANFN As String
Dim BNFPO As String
Dim sQry            As String

Dim oRecordSet01    As SAPbobsCOM.Recordset
Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

'���� ����ó��
hOpen = 0
hConnection = 0
hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
Dim nFlag As Long
nFlag = INTERNET_FLAG_PASSIVE
hConnection = InternetConnect(hOpen, FTP_SERVER, INTERNET_INVALID_PORT_NUMBER, _
    FTP_USER, FTP_PASSWORD, INTERNET_SERVICE_FTP, nFlag, 0)
        
Dir = oForm01.Items("Dir").Specific.VALUE

oFileName = oForm01.Items("FileName").Specific.VALUE

pos = InStr(oFileName, ".")

FileName = oForm01.Items("E_BANFN").Specific.VALUE + oForm01.Items("E_BNFPO").Specific.VALUE + Trim(Mid(oFileName, pos, 10))


        
'UpLoad Dir, "ftp/trans/group/" & FileName, ASCII
UpLoad Dir, "trans/group/" & FileName, ASCII

'���� ����ó��

If hConnection <> 0 Then InternetCloseHandle (hConnection)
If hOpen <> 0 Then InternetCloseHandle (hOpen)
hConnection = 0
hOpen = 0

'R3 ÷������ ���� ����

Set oSapConnection01 = CreateObject("SAP.Functions")
oSapConnection01.Connection.User = "ifuser"
oSapConnection01.Connection.Password = "pdauser"
'oSapConnection01.Connection.Client = "710"  '//�׽�Ʈ
oSapConnection01.Connection.Client = "210"  '// �
'oSapConnection01.Connection.ApplicationServer = "192.1.11.7"  '�׽�Ʈ����
oSapConnection01.Connection.ApplicationServer = "192.1.1.217"
oSapConnection01.Connection.language = "KO"
oSapConnection01.Connection.SystemNumber = "01"
'    oSapConnection01.Connection.SystemNumber = "00" '192.1.11.3 �϶��� "00"
If Not oSapConnection01.Connection.Logon(0, True) Then
    MDC_Com.MDC_GF_Message "R/3 ������ �����Ҽ� �����ϴ�.", "E"
    GoTo FTP_Trans_Exit
End If

Dim oFunction01 As Object
Set oFunction01 = oSapConnection01.Add("ZMM_INTF_GROUP_FILE")

If oFunction01 Is Nothing Then
    MDC_Com.MDC_GF_Message "�Լ�(ZMM_INTF_GROUP_FILE) ��������.", "E"
    GoTo FTP_Trans_Exit
End If

BANFN = oForm01.Items("E_BANFN").Specific.VALUE
BNFPO = oForm01.Items("E_BNFPO").Specific.VALUE

oFunction01.Exports("I_BANFN") = BANFN
oFunction01.Exports("I_BNFPO") = BNFPO
oFunction01.Exports("I_FILENAME") = FileName 'oForm01.Items("FileName").Specific.VALUE


If Not oFunction01.Call Then
    MDC_Com.MDC_GF_Message "���ձ���(R/3) �Լ�(ZMM_INTF_GROUP_FILE)ȣ���� �����߻�", "E"
    GoTo FTP_Trans_Exit
Else
    If (oFunction01.Imports("E_MESSAGE").VALUE <> "") Then '//�����޽���
'        ReturnValue = oFunction01.Imports("E_BANFN").VALUE ' & "/" & oFunction01.Imports("E_BNFPO").VALUE '//���ձ��ſ�û��ȣ '//���ձ��ſ�û ǰ���ȣ
        ReturnValue = oFunction01.Imports("E_MESSAGE").VALUE
        oForm01.Items("ReturnMsg").Specific.VALUE = ReturnValue
        
'        sQry = " Update [@PS_MM010L] "
'        sQry = sQry + " Set U_FILENAME = '" & oFileName & "'"
'        sQry = sQry + " Where U_E_BANFN = '" & BANFN & "'"
'        sQry = sQry + " And U_E_BNFPO = '" & BNFPO & "'"
'
'        oRecordSet01.DoQuery sQry
        
    Else
        sQry = " Update [@PS_MM010L] "
        sQry = sQry + " Set U_FILENAME = '" & oFileName & "'"
        sQry = sQry + " Where U_E_BANFN = '" & BANFN & "'"
        sQry = sQry + " And U_E_BNFPO = '" & BNFPO & "'"
        
        oRecordSet01.DoQuery sQry
        
'        Call MDC_Com.MDC_GF_Message(oFunction01.Imports("E_MESSAGE").VALUE, "E")
'        GoTo FTP_Trans_Exit
    End If
End If

'MsgBox "���ε� ����!"


Exit Sub
FTP_Trans_Exit:
    If Not (oSapConnection01.Connection Is Nothing) Then
        oSapConnection01.Connection.logoff
        Set oSapConnection01 = Nothing
    End If

    Set oFunction01 = Nothing
    
    '//��������
    If hConnection <> 0 Then InternetCloseHandle (hConnection)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    hConnection = 0
    hOpen = 0
    
    Exit Sub

FTP_Trans_Error:
    
   
End Sub
Private Sub FileSearch()
On Error GoTo FileSearch_Error
   
Dim sFile       As String
Dim pos As Integer

Dim FileName As String


sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.*", "���ϼ���", "C:\")
If sFile = "" Then
        Sbo_Application.StatusBar.SetText "������ ������ �ּ���.", bmt_Short, smt_Error
        Exit Sub
    Else
'        If Mid(Right(sFile, 4), 1, 3) = "pdf" Or Mid(Right(sFile, 4), 1, 3) = "hwp" Or Mid(Right(sFile, 4), 1, 3) = "ppt" Or Mid(Right(sFile, 5), 1, 4) = "pptx" Then
            oForm01.Items("Dir").Specific.VALUE = sFile
            
            pos = InStr(sFile, "\")
            
            FileName = Trim(Mid(sFile, pos + 1, 100))
            
            pos = InStr(FileName, "\")
            If pos > 0 Then
                FileName = Trim(Mid(FileName, pos + 1, 100))
                pos = InStr(FileName, "\")
                If pos > 0 Then
                    FileName = Trim(Mid(FileName, pos + 1, 100))
                    pos = InStr(FileName, "\")
                    If pos > 0 Then
                        FileName = Trim(Mid(FileName, pos + 1, 100))
                        pos = InStr(FileName, "\")
                        If pos > 0 Then
                            FileName = Trim(Mid(FileName, pos + 1, 100))
                            pos = InStr(FileName, "\")
                        Else
                            oForm01.Items("FileName").Specific.VALUE = FileName
                        End If
                    Else
                        oForm01.Items("FileName").Specific.VALUE = FileName
                    End If
                Else
                    oForm01.Items("FileName").Specific.VALUE = FileName
                End If
            Else
                oForm01.Items("FileName").Specific.VALUE = FileName
            End If
            

            
            
'            pos = InStr(FileName, "\")
'            If pos = 0 Then
'                oForm01.Items("FileName").Specific.VALUE = FileName
'            Else
'                FileName = Trim(Mid(FileName, pos + 1, 100))
'                pos = InStr(FileName, "\")
'                If pos = 0 Then
'                    oForm01.Items("FileName").Specific.VALUE = FileName
'                Else
'                    FileName = Trim(Mid(FileName, pos + 1, 100))
'                    pos = InStr(FileName, "\")
'                End If
'            End If
            
            
            
'        Else
'            Sbo_Application.StatusBar.SetText "pdf������ �ƴմϴ�.", bmt_Short, smt_Error
'            Exit Sub
'        End If
    End If

Exit Sub
FileSearch_Error:
    
   
End Sub

Sub UpLoad(szFileLocal As String, szFileRemote As String, dwType As Long)
    Dim bRet As Boolean
    bRet = FtpPutFile(hConnection, szFileLocal, szFileRemote, dwType, 0)
    'MDC_Com.MDC_GF_Message String(bRet), "E"
    
End Sub

Sub DownLoad(szFileRemote As String, szFileLocal As String, dwType As Long)
    Dim bRet As Boolean
    bRet = FtpGetFile(hConnection, szFileRemote, szFileLocal, False, INTERNET_FLAG_RELOAD, dwType, 0)
End Sub

Private Sub cmdQuit_Click()
    '���� ����ó��
    If hConnection <> 0 Then InternetCloseHandle (hConnection)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    hConnection = 0
    hOpen = 0
    End
End Sub

Private Sub cmdUpLoad_Click()
    UpLoad App.Path & "\test.txt", "public_html/test.txt", ASCII
    MsgBox "���ε� ����!"
End Sub

Private Sub cmdDownLoad_Click()
    DownLoad "public_html/test.txt", App.Path & "\test.txt", ASCII
    MsgBox "�ٿ�ε� ����!"
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdQuit_Click
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////





