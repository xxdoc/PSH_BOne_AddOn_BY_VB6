VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP778"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP778.cls
'//  Module         : PP
'//  Description    : 휘팅이동등록
'//  FormType       : PS_PP778
'//  Create Date    : 2010.10.21
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_PP778H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_PP778L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP778.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP778_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call FormClear
    Call Make_MovDocNo
    Call Matrix_AddRow(1, 0, True) 'oMat01
    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
        
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
    '//디비데이터 소스 개체 할당
    Set oDS_PS_PP778H = oForm01.DataSources.DBDataSources("@PS_PP778H")
    Set oDS_PS_PP778L = oForm01.DataSources.DBDataSources("@PS_PP778L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oDS_PS_PP778H.setValue "U_RegiDate", 0, Format(Now, "YYYYMMDD")
    
    '담당자
    oDS_PS_PP778H.setValue "U_CntcCode", 0, User_MSTCOD()
    FlushToItemValue "CntcCode"
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
'    Call oCombo.Select(User_BPLId(), psk_ByValue)
    Call oCombo.Select(3, psk_Index)
    
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
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                
                ElseIf pval.ItemUID = "Print" Then '출력버튼 클릭시
                    Call Print_Query
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    '//라인
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "PP777HL" Then
                            If oMat01.Columns("PP777HL").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
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
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                        FormItemEnabled
                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "CntcCode" Or pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "RegiDate" Then
                        Call Make_MovDocNo
                    End If
                    '//라인
                    If pval.ItemUID = "Mat01" And (pval.ColUID = "PP777HL") Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                    End If
                End If
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
                Set oMat01 = Nothing
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
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1285": '복원
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1285": '복원
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
                    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineId").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
    
                    oMat01.FlushToDataSource
                    oDS_PS_PP778L.RemoveRecord oDS_PS_PP778L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
            
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                
            Case "1282": '추가
                Call FormItemEnabled
                Call FormClear
                oDS_PS_PP778H.setValue "U_RegiDate", 0, Format(Now, "YYYYMMDD")
                Call Matrix_AddRow(1, 0, True) 'oMat01
                Call Make_MovDocNo
                
                Dim oCombo          As SAPbouiCOM.ComboBox
                Set oCombo = oForm01.Items("BPLId").Specific
                oForm01.Items("BPLId").Enabled = True
                'Call oCombo.Select(User_BPLId(), psk_Index)
                Call oCombo.Select(3, psk_Index)
                
                oDS_PS_PP778H.setValue "U_CntcCode", 0, User_MSTCOD()
                FlushToItemValue "CntcCode"
                
                oForm01.Items("CntcCode").Click ct_Regular
                oForm01.Items("BPLId").Enabled = False

            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("PP777HL").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_PP778H.GetValue("Status", 0) = "O" Then
                            Matrix_AddRow 1, oMat01.RowCount, False
                        End If
                    End If
                End If
  
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
        '//작업
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

Private Sub FormItemEnabled()
    If oForm01.Mode = fm_FIND_MODE Then
        oForm01.Items("DocNum").Enabled = True
'        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DeliArea").Enabled = True
        oForm01.Items("CarNo").Enabled = True
        oForm01.Items("TransCom").Enabled = True
        oForm01.Items("Fee").Enabled = True
        oForm01.Items("RegiDate").Enabled = True
        oMat01.Columns("PP777HL").Editable = True
        
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("DocNum").Enabled = False
'        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DeliArea").Enabled = True
        oForm01.Items("CarNo").Enabled = True
        oForm01.Items("TransCom").Enabled = True
        oForm01.Items("Fee").Enabled = True
        oForm01.Items("RegiDate").Enabled = True
        oMat01.Columns("PP777HL").Editable = True
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        If oForm01.Items("Canceled").Specific.VALUE = "Y" Then
            oForm01.Items("Comments").Click ct_Regular
            oForm01.Items("DocNum").Enabled = False
'            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("CntcCode").Enabled = False
            oForm01.Items("CardCode").Enabled = False
            oForm01.Items("DeliArea").Enabled = False
            oForm01.Items("CarNo").Enabled = False
            oForm01.Items("TransCom").Enabled = False
            oForm01.Items("Fee").Enabled = False
            oForm01.Items("RegiDate").Enabled = False
            oMat01.Columns("PP777HL").Editable = False
        Else
            oForm01.Items("DocNum").Enabled = False
'            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("DeliArea").Enabled = True
            oForm01.Items("CarNo").Enabled = True
            oForm01.Items("TransCom").Enabled = True
            oForm01.Items("Fee").Enabled = True
            oForm01.Items("RegiDate").Enabled = True
            oMat01.Columns("PP777HL").Editable = False
        End If
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP778'", "")
    If DocNum = 0 Then
        oForm01.Items("DocNum").Specific.VALUE = 1
    Else
        oForm01.Items("DocNum").Specific.VALUE = DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oMat As Integer, ByVal oRow As Long, Optional Insert_YN As Boolean)
On Error GoTo Matrix_AddRow_Error
        
    Select Case oMat
        Case 1: 'oMat01
            If Insert_YN = False Then
                oRow = oMat01.RowCount
                oDS_PS_PP778L.InsertRecord (oRow)
            End If
            '수입내역
            oDS_PS_PP778L.Offset = oRow
            oDS_PS_PP778L.setValue "LineId", oRow, oRow + 1
            oDS_PS_PP778L.setValue "U_PP777HL", oRow, ""
            oDS_PS_PP778L.setValue "U_ItemCode", oRow, ""
            oDS_PS_PP778L.setValue "U_ItemName", oRow, ""
            oDS_PS_PP778L.setValue "U_Size", oRow, ""
            oDS_PS_PP778L.setValue "U_Mark", oRow, ""
            oDS_PS_PP778L.setValue "U_Qty", oRow, ""
            oDS_PS_PP778L.setValue "U_Weight", oRow, ""
            oDS_PS_PP778L.setValue "U_DocDate", oRow, ""
            oMat01.LoadFromDataSource

'        Case 2: 'oMat02
'            If Insert_YN = False Then
'                oRow = oMat2.RowCount
'                oDS_ZPP140M.InsertRecord (oRow)
'            End If
'            '수출내역
'            oDS_ZPP140M.Offset = oRow
'            oDS_ZPP140M.setValue "LineId", oRow, oRow + 1
'            oDS_ZPP140M.setValue "U_ConfDate", oRow, ""
'            oDS_ZPP140M.setValue "U_ConfNo", oRow, ""
'            oDS_ZPP140M.setValue "U_Size", oRow, ""
'            oDS_ZPP140M.setValue "U_ExpQty", oRow, ""
'            oDS_ZPP140M.setValue "U_RfndQty", oRow, ""
'            oMat02.LoadFromDataSource
        
    End Select
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Matrix_AddRow_Error:
    MDC_Com.MDC_GF_Message "Matrix_AddRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordset    As SAPbobsCOM.Recordset
        
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "CntcCode"
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_PP778H.GetValue("U_CntcCode", 0)) & "'"
            oRecordset.DoQuery sQry
            oDS_PS_PP778H.setValue "U_CntcName", 0, Trim(oRecordset.Fields(0).VALUE)
        Case "CardCode"
            sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" & Trim(oDS_PS_PP778H.GetValue("U_CardCode", 0)) & "'"
            oRecordset.DoQuery sQry
            oDS_PS_PP778H.setValue "U_CardName", 0, Trim(oRecordset.Fields(0).VALUE)
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
            Case "PP777HL"
                oMat01.FlushToDataSource
                oDS_PS_PP778L.Offset = oRow - 1
                'oMat01.SetLineData oRow
                
                i = InStr(oMat01.Columns("PP777HL").Cells(oRow).Specific.String, "-")
                
                DocNum = Left(oMat01.Columns("PP777HL").Cells(oRow).Specific.String, i - 1)
                LineId = Mid(oMat01.Columns("PP777HL").Cells(oRow).Specific.String, i + 1)
                
                sQry = "select b.U_ItemCode, b.U_ItemName, c.U_Size, c.U_Mark, d.name, b.U_PkQty, b.U_PkWt, a.U_DocDate "
                sQry = sQry & "from [@PS_PP777H] a inner join [@PS_PP777L] b on a.docentry=b.docentry "
                sQry = sQry & "left  join OITM c on b.U_ItemCode=c.ItemCode "
                sQry = sQry & "left  join [@PSH_MARK] d on c.U_Mark=d.Code "
                sQry = sQry & "Where a.DocNum = '" & DocNum & "'"
                sQry = sQry & "and b.LineId = '" & LineId & "'"
                oRecordset.DoQuery sQry
                
                oDS_PS_PP778L.setValue "U_ItemCode", oRow - 1, oRecordset.Fields("U_ItemCode").VALUE
                oDS_PS_PP778L.setValue "U_ItemName", oRow - 1, oRecordset.Fields("U_ItemName").VALUE
                oDS_PS_PP778L.setValue "U_Size", oRow - 1, oRecordset.Fields("U_Size").VALUE
                oDS_PS_PP778L.setValue "U_Mark", oRow - 1, oRecordset.Fields("name").VALUE
                oDS_PS_PP778L.setValue "U_Qty", oRow - 1, oRecordset.Fields("U_PkQty").VALUE
                oDS_PS_PP778L.setValue "U_Weight", oRow - 1, oRecordset.Fields("U_PkWt").VALUE
                oDS_PS_PP778L.setValue "U_DocDate", oRow - 1, Format(oRecordset.Fields("U_DocDate").VALUE, "YYYYMMDD")
                oMat01.SetLineData oRow
                
                '--------------------------------------------------------------------------------------------
                If oRow = oMat01.RowCount And Trim(oDS_PS_PP778L.GetValue("U_PP777HL", oRow - 1)) <> "" Then
                '// 다음 라인 추가
                    Call Matrix_AddRow(1, 0, False)
                    oMat01.Columns("PP777HL").Cells(oRow).Click ct_Regular
                End If

         End Select
    End If
     
    Set oRecordset = Nothing
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
        Case Trim(oDS_PS_PP778H.GetValue("U_BPLId", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP778H.GetValue("U_CntcCode", 0)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP778H.GetValue("U_CntcName", 0)) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_PP778H.GetValue("U_CardCode", 0)) = ""
'            ErrNum = 4
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_PP778H.GetValue("U_CardName", 0)) = ""
'            ErrNum = 5
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_PP778H.GetValue("U_CarNo", 0)) = ""
'            ErrNum = 6
'            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_PP778H.GetValue("U_Fee", 0)) = "" Or Trim(oDS_PS_PP778H.GetValue("U_Fee", 0)) = 0
'            ErrNum = 7
'            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP778H.GetValue("U_RegiDate", 0)) = ""
            ErrNum = 8
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP778H.GetValue("U_MovDocNo", 0)) = ""
            ErrNum = 9
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "담당자는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "담장자명이 없습니다. 담당자코드를 확인하여 주십시오.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "납품처는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "납품처명이 없습니다. 납품처코드를 확인하여 주십시오.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "차량번호는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "운임이 없거나 값이 0 입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "등록일자는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "이동등록번호는 필수사항입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 1 Then '// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
            oDS_PS_PP778L.Offset = i
            If Trim(oDS_PS_PP778L.GetValue("U_PP777HL", i)) = "" Then
                ErrNum = 2
                oMat01.Columns("PP777HL").Cells(i + 1).Click ct_Regular
                GoTo MatrixSpaceLineDel_Error
            End If
        Next i
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        oDS_PS_PP778L.RemoveRecord oDS_PS_PP778L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource

    Set oRecordset = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordset = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "벌크반품문서 번호는 필수입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub Make_MovDocNo() '이동등록번호생성
On Error GoTo Make_MovDocNo_Error
    
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)

    '/ Procedure 실행
    sQry = "EXEC PS_PP778_01 '" & oDS_PS_PP778H.GetValue("U_RegiDate", 0) & "'"
    oRecordset.DoQuery sQry

    oDS_PS_PP778H.setValue "U_MovDocNo", 0, Trim(oRecordset.Fields(0).VALUE)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Make_MovDocNo_Error:
    MDC_Com.MDC_GF_Message "Make_MovDocNo_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Print_Query()
On Error GoTo Print_Query_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry        As String
    Dim DocNum          As Integer
    Dim oRecordset      As SAPbobsCOM.Recordset
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// 조회조건문
    DocNum = Trim(oForm01.Items("DocNum").Specific.VALUE)

    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP778] 출고원부/반출증"
    ReportName = "PS_PP778_01.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
'    gRpt_Formula(1) = "YYYY"
'    gRpt_Formula_Value(1) = Left(YearMon, 4)
'    gRpt_Formula(2) = "MM"
'    gRpt_Formula_Value(2) = Right(YearMon, 2)
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행
    sQry = "EXEC [PS_PP778_02] '" & DocNum & "'"
    oRecordset.DoQuery sQry
    If oRecordset.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordset = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordset = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
