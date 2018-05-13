VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM007.cls
'//  Module         : MM
'//  Description    : 품목마스터 생성요청
'//  FormType       : PS_MM007
'//  Create Date    : 2015.05.06
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************


Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_MM007H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM007L As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oSeq&
Private oCount&



'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional sDocNum As String, Optional SBPLID As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM007.srf")
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
    
    oFormUniqueID01 = "PS_MM007_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
'    oForm01.Mode = fm_ADD_MODE
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    If sDocNum <> "" Then
        oForm01.Mode = fm_FIND_MODE
    Else
        oForm01.Mode = fm_ADD_MODE
    End If
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
'    Call Initialization
    Call FormClear
    Call Add_MatrixRow(0, True)
        
'    FormItemEnabled
    
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), True        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    
    oForm01.Update
    oForm01.Freeze False
    
    If sDocNum <> "" Then

        Call FormItemEnabled

        oForm01.Items("DocNum").Specific.VALUE = sDocNum
'        Call oForm01.Items("BPLId").Specific.Select(SBPLID, psk_ByValue)
                
        oForm01.Items("CntcCode").Specific.VALUE = ""
        
        Call oForm01.Items("1").Click(ct_Regular)
        
    Else
    
        Call Initialization
        
    End If

    
    
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
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Dim SumWeight As Currency, SumQty As Long
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$, BPLID$
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FileName As String
    
    Dim Path As String
    Dim exeFileName As String
    
    Dim CntcCode As String
    Dim TelNo As String
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                            
                        End If
                        
                        '//통합구매
                        If oForm01.Mode = fm_UPDATE_MODE And Trim(oDS_PS_MM007H.GetValue("U_RFCAdms", 0)) = "Y" Then '//신규코드 생성요청
                            Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("통합구매 신규코드생성 요청!", oMat01.VisualRowCount - 2 + 1, False)
                            
                            ErrNum = 101
                            oSeq = 1
                            oCount = 1
                            Seq = 1
                            oMat01.FlushToDataSource
                            
                            
                            DueDate = Trim(oDS_PS_MM007H.GetValue("U_DocDate", 0))
                            BPLID = Trim(oDS_PS_MM007H.GetValue("U_BPLId", 0))
                            
                            Dim CntcName As String
                            Dim OffcTel As String
                            Dim TeamName As String
                            Dim EMail As String
                            Dim MATNR As String
                            Dim MAKTX As String
                            Dim MTART As String
                            Dim MEINS As String
                            Dim MATKL As String
                            Dim ZUSE As String   '//용도
                            Dim NORMT As String  '//메이커
                            
                            '//담당자 성명(전화번호)
                            CntcName = oDS_PS_MM007H.GetValue("U_CntcName", 0)
                            OffcTel = oDS_PS_MM007H.GetValue("U_OffcTel", 0)
                            EMail = oDS_PS_MM007H.GetValue("U_EMail", 0)
                            
                            For i = 0 To oMat01.VisualRowCount - 2
                                If Trim(oDS_PS_MM007L.GetValue("U_TransYN", i)) = "N" Or Trim(oDS_PS_MM007L.GetValue("U_TransYN", i)) = "" Then
                                    RequestDate = Trim(oDS_PS_MM007L.GetValue("U_ReqDate", 0)) '//요청일자
                                    RequestNo = Trim(oDS_PS_MM007L.GetValue("U_ReqNo", i)) '//요청번호
                                    
                                    If Trim(oDS_PS_MM007L.GetValue("U_RspName", i)) = "" Then
                                        TeamName = Trim(oDS_PS_MM007L.GetValue("U_TeamName", i))
                                    Else
                                        TeamName = Trim(oDS_PS_MM007L.GetValue("U_TeamName", i)) + " / " + Trim(oDS_PS_MM007L.GetValue("U_RspName", i)) '//부서/담당명
                                    End If
                                    MATNR = Trim(oDS_PS_MM007L.GetValue("U_MATNR", i))  '//상품번호
                                    MAKTX = Trim(oDS_PS_MM007L.GetValue("U_MAKTX", i)) '//상품내역
                                    MTART = Trim(oDS_PS_MM007L.GetValue("U_MTART", i)) '//상품유형
                                    MEINS = Trim(oDS_PS_MM007L.GetValue("U_MEINS", i)) '//기본단위
                                    MATKL = Trim(oDS_PS_MM007L.GetValue("U_MATKL", i)) '//상품범주
                                    ZUSE = Trim(oDS_PS_MM007L.GetValue("U_ZUSE", i)) '//용도
                                    NORMT = Trim(oDS_PS_MM007L.GetValue("U_NORMT", i)) '//메이커(MAKER)
                                    
                                    
                                    '///////////////////////////////////
                                    '//사업장, 요청번호, 팀/담당명, 요청일자, 전화번호, 이메일, 상품번호(7자리), 상품내역, 상품유형, 기본단위, 상품범주, 메이커,
                                    RFC_Sender = Trim(PS_RFC_Sender(BPLID, RequestNo, CntcName, TeamName, RequestDate, OffcTel, EMail, MATNR, MAKTX, MTART, MEINS, MATKL, ZUSE, NORMT, i, oMat01.VisualRowCount - 2))
                                    
                                    If RFC_Sender <> "" Then
                                        oDS_PS_MM007L.setValue "U_TransYN", i, Trim(RFC_Sender)
                                    End If
                                    
'
                                    
                                    ProgressBar01.VALUE = ProgressBar01.VALUE + 1
                                    ProgressBar01.Text = ProgressBar01.VALUE & "/" & oMat01.VisualRowCount - 2 + 1 & "건 처리중...!"
                                End If
                            Next
                            oMat01.LoadFromDataSource
                            
                            
                            ErrNum = 0
                            
                            ProgressBar01.Stop
                            Set ProgressBar01 = Nothing
                        End If
                        Call Delete_EmptyRow
                        oLast_Mode = oForm01.Mode
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        oLast_Mode = oForm01.Mode
                    End If
                End If
                If pval.ItemUID = "Btn_01" Then
                    Path = "\\191.1.1.221\ERP_Project\9999. Add_On\인터페이스05(품목생성요청)\"
                    exeFileName = "\\191.1.1.221\ERP_Project\9999. Add_On\인터페이스05(품목생성요청)\RFC_INTERFACE_OITM.exe"
                    Call ShellExecute(0, "Open", exeFileName, "", Path, 5)
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "TeamCode" Then
                            If oMat01.Columns("TeamCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "RspCode" Then
                            If oMat01.Columns("RspCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "MATKL" Then
                            If oMat01.Columns("MATKL").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "MATNR" Then '//품목코드를 사용안함
                            If oMat01.Columns("MATNR").Cells(pval.Row).Specific.VALUE = "" Then
                                Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                                BubbleEvent = False
                            End If
                        End If

'                        If pval.ColUID = "MAKTX" Then
'                            If oMat01.Columns("MAKTX").Cells(pval.Row).Specific.VALUE = "" Then
'                                Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
'                                BubbleEvent = False
'                            End If
'                        End If
                        
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'                If pval.ItemUID = "Mat01" Then
'                   If pval.ColUID = "E_BANFN" Then
'                        Dim TempForm01 As Variant
'                        '통합구매이며 통합구매요청번호가 있을때
'                        If oForm01.Items("PQType").Specific.VALUE = "20" And oMat01.Columns("E_BANFN").Cells(pval.Row).Specific.String <> "" Then
'                            Set TempForm01 = New PS_FTP
'                            Call TempForm01.LoadForm(oMat01.Columns("CGNo").Cells(pval.Row).Specific.String, oMat01.Columns("E_BANFN").Cells(pval.Row).Specific.String, oMat01.Columns("E_BNFPO").Cells(pval.Row).Specific.String)
'                            BubbleEvent = False
'                        Else
'
'                        End If
'                   End If
'
'                   If pval.ColUID = "CGNo" Then
'                        Dim TempForm02 As Variant
'                        If oMat01.Columns("CGNo").Cells(pval.Row).Specific.String <> "" Then
'                            Set TempForm02 = New PS_MM007_S
'                            Call TempForm02.LoadForm(oMat01.Columns("CGNo").Cells(pval.Row).Specific.String)
'                            BubbleEvent = False
'                        Else
'
'                        End If
'                   End If
'
'                End If
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
                If pval.ItemUID = "1" Then
                    
                    If oForm01.Mode = fm_OK_MODE Then
                        If oLast_Mode = fm_UPDATE_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            oLast_Mode = 100
                        ElseIf oLast_Mode = fm_FIND_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            FormItemEnabled
                            oLast_Mode = 100
                        End If
                    ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "ItmsGrp" Or pval.ItemUID = "BPLId" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_MM007L.Clear
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call Add_MatrixRow(0, False)
                    End If
                    
                    
                    oForm01.Freeze False
                End If
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then
        
                    If pval.Row > 0 Then
                    
                        oLast_Item_UID = pval.ItemUID
                        oLast_Col_UID = pval.ColUID
                        oLast_Col_Row = pval.Row
                        
                        Call oMat01.SelectRow(pval.Row, True, False)
                        
                    End If
                    
                End If
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CntcCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "TeamCode" Or pval.ColUID = "RspCode" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "MATNR" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "MATKL" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_MM007H = Nothing
                Set oDS_PS_MM007L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    oForm01.Freeze False
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                If oForm01.Items("RFCAdms").Specific.VALUE = "Y" Then
                    MDC_Com.MDC_GF_Message "통합구매 코드생성한 자료입니다. 취소할 수 없습니다.", "E"
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
'[1284:취소] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1284": '취소
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_MM007L.RemoveRecord oDS_PS_MM007L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                
                '//아이디별 사업장 세팅
                oForm01.Items("BPLId").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
                '//아이디별 사번 세팅
                If MDC_PS_Common.User_SuperUserYN = "Y" Then '슈퍼유저인경우는 사번 추가 하지 않음(2018.04.02 송명규)
                    oForm01.Items("CntcCode").Specific.VALUE = ""
                Else
                    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
                End If
                
                oForm01.Freeze False
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                oDS_PS_MM007H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                Call Add_MatrixRow(0, True)
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("ReqNo").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_MM007H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                End If
        End Select
    End If
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
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
        '//작업
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_MM007H = oForm01.DataSources.DBDataSources("@PS_MM007H")
    Set oDS_PS_MM007L = oForm01.DataSources.DBDataSources("@PS_MM007L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    
    oDS_PS_MM007H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
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
    sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//품목그룹
    Set oCombo = oForm01.Items("ItmsGrp").Specific
    oCombo.ValidValues.Add "105", "[5]저장품"
    Call oDS_PS_MM007H.setValue("U_ItmsGrp", 0, "105")
    'Call oCombo.Select("0", psk_Index)
    
    Set oCombo = oForm01.Items("RFCAdms").Specific
    oCombo.ValidValues.Add "N", "미승인"
    oCombo.ValidValues.Add "Y", "승인"
    Call oDS_PS_MM007H.setValue("U_RFCAdms", 0, "N")
    
    
    '//자재유형
'    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oMat01.Columns("MTART").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
    
    '//기본단위
'    sQry = "SELECT Name, Name From [@PSH_UOMORG] Order by Code"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oMat01.Columns("MEINS").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
    
    '//상품범주
'    sQry = "SELECT Code, Name From [@PSH_UOMORG] Order by Code"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oMat01.Columns("MATKL").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
    
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
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
        
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CF_ChooseFromList_Error:
    MDC_Com.MDC_GF_Message "CF_ChooseFromList_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error

    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("ItmsGrp").Enabled = True
        oForm01.Items("RFCAdms").Enabled = True
        oForm01.Items("DocDate").Enabled = True
                
        oForm01.Items("Mat01").Enabled = True
        
        '//품목그룹
        Call oDS_PS_MM007H.setValue("U_ItmsGrp", 0, "105")
        '//승인여부
        Call oDS_PS_MM007H.setValue("U_RFCAdms", 0, "N")
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("ItmsGrp").Enabled = True
        oForm01.Items("OffcTel").Enabled = True
        oForm01.Items("RFCAdms").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        
        oForm01.Items("Mat01").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        
        If Trim(oDS_PS_MM007H.GetValue("U_RFCAdms", 0)) = "Y" Or Trim(oDS_PS_MM007H.GetValue("Canceled", 0)) = "Y" Then
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("CntcCode").Enabled = False
            oForm01.Items("ItmsGrp").Enabled = False
            oForm01.Items("OffcTel").Enabled = False
            oForm01.Items("RFCAdms").Enabled = False
            oForm01.Items("DocDate").Enabled = False
        Else
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("ItmsGrp").Enabled = True
            oForm01.Items("OffcTel").Enabled = True
            oForm01.Items("RFCAdms").Enabled = True
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("Mat01").Enabled = True
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
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM007'", "")
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

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM007L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM007L.offset = oRow
    oDS_PS_MM007L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim SumWeight As Currency, SumQty As Long
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CntcCode"
            sQry = "Select U_FullName, U_OffcTel, U_eMail From [@PH_PY001A] Where Code = '" & Trim(oDS_PS_MM007H.GetValue("U_CntcCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM007H.setValue "U_CntcName", 0, Trim(oRecordSet01.Fields(0).VALUE)
            oDS_PS_MM007H.setValue "U_OffcTel", 0, Trim(oRecordSet01.Fields(1).VALUE)
            oDS_PS_MM007H.setValue "U_EMail", 0, Trim(oRecordSet01.Fields(2).VALUE)
        Case "Mat01"
            If oCol = "TeamCode" Then '//부서
                oForm01.Freeze True
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("TeamCode").Cells(oRow).Specific.VALUE) <> "" Then
                    oMat01.FlushToDataSource
                    Add_MatrixRow oMat01.RowCount, False
                    oMat01.Columns("TeamCode").Cells(oRow).Click ct_Regular
                End If
                
                sQry = "Select U_CodeNm From [@PS_HR200L] Where Code = '1' and U_Code = '" & Trim(oMat01.Columns("TeamCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("TeamName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CodeNm").VALUE)
                
                
                oForm01.Freeze False
            
            ElseIf oCol = "RspCode" Then '//담당
                oForm01.Freeze True
                sQry = "Select U_CodeNm From [@PS_HR200L] Where Code = '2' and U_Code = '" & Trim(oMat01.Columns("RspCode").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("RspName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_CodeNm").VALUE)
                oForm01.Freeze False
            ElseIf oCol = "MATKL" Then '//품목분류
                oForm01.Freeze True
                sQry = "Select CodeName = U_CodeName From [@PSH_ITMMSORT] Where U_Code = '" & Trim(oMat01.Columns("MATKL").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("MATKLNM").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("CodeName").VALUE)
                oForm01.Freeze False
            End If
            
            
    End Select

'    Set oRecordset = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_MM007H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM007H.GetValue("U_CntcCode", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM007H.GetValue("U_OffcTel", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM007H.GetValue("U_ItmsGrp", 0) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM007H.GetValue("U_DocDate", 0) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "담당자는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "전화번호는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "품목그룹은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "전기일은 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ReqDate         As String
    Dim ReqNo           As String
    Dim Seq             As String
    Dim DocDate         As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 And oDS_PS_MM007L.GetValue("U_ReqNo", 0) = "" Then
        ErrNum = 2
        GoTo MatrixSpaceLineDel_Error
    End If
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    sQry = "Select ReqNo = Max(U_ReqNo) From [@PS_MM007L] Where Left(U_ReqNo,8) = '" & DocDate & "'"
    oRecordSet01.DoQuery sQry
    
    If Trim(oRecordSet01.Fields("ReqNo").VALUE) = "" Then
        ReqNo = DocDate + "000"
        Seq = "000"
    Else
        ReqNo = Trim(oRecordSet01.Fields("ReqNo").VALUE)
        Seq = Right(Trim(oRecordSet01.Fields("ReqNo").VALUE), 3)
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        oMat01.Columns("ReqDate").Cells(i).Specific.VALUE = DocDate
        If Trim(oMat01.Columns("ReqNo").Cells(i).Specific.VALUE) = "" Then
            oMat01.Columns("ReqNo").Cells(i).Specific.VALUE = Left(ReqNo, 8) + Right("000" & Trim(Str(Val(Seq) + i)), 3)
        End If
    Next
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_MM007L.GetValue("U_ReqNo", i) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM007L.GetValue("U_TeamName", i) = ""
                ErrNum = 4
                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM007L.GetValue("U_RspName", i) = ""
'                MDC_Com.MDC_GF_Message "담당이 없습니다. 확인하시고 진행하세요.", "E"
            Case oDS_PS_MM007L.GetValue("U_MATNR", i) = ""
                ErrNum = 6
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM007L.GetValue("U_MAKTX", i) = ""
                ErrNum = 7
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM007L.GetValue("U_NORMT", i) = ""
                ErrNum = 8
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM007L.GetValue("U_MEINS", i) = ""
                ErrNum = 9
            Case Len(Trim(oDS_PS_MM007L.GetValue("U_MEINS", i))) > "2"
                ErrNum = 11
                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM007L.GetValue("U_MATKL", i) = ""
                ErrNum = 10
                GoTo MatrixSpaceLineDel_Error
        End Select
    Next
    
   
    oMat01.LoadFromDataSource

    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Or ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "청구번호는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "팀명은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "상품번호 7자리는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "자재내역은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "메이커는 필수사항입니다. 없을시'국산'이라고 입력바랍니다.", "E"
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "기본단위는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "상품범주는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "기본단위는 두 글자로 입력하셔야 합니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_MM007L.GetValue("U_ReqNo", i)) = "" Then
            oDS_PS_MM007L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Function PS_RFC_Sender(ByVal BPLID As String, ByVal RequestNo As String, ByVal CntcName As String, ByVal TeamName As String, ByVal ReqDate As String, ByVal OffcTel As String, ByVal EMail As String, ByVal MATNR As String, ByVal MAKTX As String, ByVal MTART As String, ByVal MEINS As String, ByVal MATKL As String, ByVal ZUSE As String, ByVal NORMT As String, ByVal i&, ByVal LastRow&) As String
On Error GoTo RFC_Sender_Error

    Dim ReturnValue As String
    Dim WERKS As String
    Dim Rotate As String
    
    If oCount = 1 Then
        Set oSapConnection01 = CreateObject("SAP.Functions")
        oSapConnection01.Connection.User = "ifuser"
        oSapConnection01.Connection.Password = "pdauser"
'        oSapConnection01.Connection.Client = "710"   '//테스트
        oSapConnection01.Connection.Client = "210"  '//운영
'        oSapConnection01.Connection.ApplicationServer = "192.1.11.7"  '//테스트
'        oSapConnection01.Connection.ApplicationServer = "192.1.11.3"   '//운영(디버깅용)
        oSapConnection01.Connection.ApplicationServer = "192.1.1.217"   '//운영
        oSapConnection01.Connection.language = "KO"
        oSapConnection01.Connection.SystemNumber = "01"
'    oSapConnection01.Connection.SystemNumber = "00" '192.1.11.3 일때는 "00"
        If Not oSapConnection01.Connection.Logon(0, True) Then
            MDC_Com.MDC_GF_Message "안강(R/3)서버에 접속할수 없습니다.", "E"
            GoTo RFC_Sender_Exit
        End If
        
        oCount = 99
    Else
        If i = 0 Then
            Set oSapConnection01 = CreateObject("SAP.Functions")
            oSapConnection01.Connection.User = "ifuser"
            oSapConnection01.Connection.Password = "pdauser"
'            oSapConnection01.Connection.Client = "710"  '//테스트
            oSapConnection01.Connection.Client = "210"   '//운영
'            oSapConnection01.Connection.ApplicationServer = "192.1.11.7"  '테스트
'            oSapConnection01.Connection.ApplicationServer = "192.1.11.3"  '운영(디버깅용)
            oSapConnection01.Connection.ApplicationServer = "192.1.1.217"  '운영
            oSapConnection01.Connection.language = "KO"
            oSapConnection01.Connection.SystemNumber = "01"
'           oSapConnection01.Connection.SystemNumber = "00" '192.1.11.3 일때는 "00"
            If Not oSapConnection01.Connection.Logon(0, True) Then
                MDC_Com.MDC_GF_Message "안강(R/3)서버에 접속할수 없습니다.", "E"
                GoTo RFC_Sender_Exit
            End If
        Else
            If oSeq = 99 Then
                Set oSapConnection01 = CreateObject("SAP.Functions")
                oSapConnection01.Connection.User = "ifuser"
                oSapConnection01.Connection.Password = "pdauser"
'                oSapConnection01.Connection.Client = "710" '//테스트
                oSapConnection01.Connection.Client = "210" '//운영
'                oSapConnection01.Connection.ApplicationServer = "192.1.11.7" '테스트
'                oSapConnection01.Connection.ApplicationServer = "192.1.11.3" '운영(디버깅용)
                oSapConnection01.Connection.ApplicationServer = "192.1.1.217" '운영
                oSapConnection01.Connection.language = "KO"
                oSapConnection01.Connection.SystemNumber = "01"
'               oSapConnection01.Connection.SystemNumber = "00" '192.1.11.3 일때는 "00"
                If Not oSapConnection01.Connection.Logon(0, True) Then
                    MDC_Com.MDC_GF_Message "안강(R/3)서버에 접속할수 없습니다.", "E"
                    GoTo RFC_Sender_Exit
                End If
            Else
            End If
        End If
    End If
    
    
    Dim oFunction01 As Object
    Set oFunction01 = oSapConnection01.Add("ZMM_SUB_REQMAT")
    
    Select Case BPLID
        Case 1:
            WERKS = "9200"
        Case 2:
            WERKS = "9300"
        Case 3:
            WERKS = "9500"
        Case 5:
            WERKS = "9600"
    End Select
        
'    If BPLId = 1 Then
'        WERKS = "9200"
'    ElseIf BPLId = 2 Then
'        WERKS = "9300"
'    Else
'        WERKS = "9200"
'    End If
    
    oFunction01.Exports("I_WERKS") = WERKS '//플랜트 홀딩스 창원 9200, 홀딩스 부산 9300, 포장사업팀 9500, 포장온산 9600
    oFunction01.Exports("I_REQNO") = RequestNo '//요청문서번호
    oFunction01.Exports("I_MATNR") = MATNR '//상품번호(7자리)
    
    oFunction01.Exports("I_MAKTX") = MAKTX '//상품내역
    oFunction01.Exports("I_WRKST") = "" '//기본자재 ?
'    oFunction01.Exports("I_BUKRS") = "PSH" '//회사코드 '사용안함
    oFunction01.Exports("I_USRNB") = CntcName '//요청생성인
    oFunction01.Exports("I_DEPRT") = TeamName '//부서 담당
    oFunction01.Exports("I_ERSDA") = ReqDate '//요청일자
    oFunction01.Exports("I_TELNO") = OffcTel '//전화번호
    
    oFunction01.Exports("I_EMAIL") = EMail '//이메일
    
    oFunction01.Exports("I_MTART") = MTART '//자재유형
    
    oFunction01.Exports("I_MEINS") = MEINS '//기본단위
    oFunction01.Exports("I_MATKL") = MATKL '//상품범주
    oFunction01.Exports("I_NORMT") = NORMT '//(메이커)MAKER
    oFunction01.Exports("I_ZUSE") = ZUSE '//용도
    oFunction01.Exports("I_ZSUPP") = "" '//첨부물표시
    oFunction01.Exports("I_FILENAME") = "" '//첨부파일이름
    
    
    
    
    If Not (oFunction01.Call) Then
        MDC_Com.MDC_GF_Message "안강(R/3)서버 함수호출중 오류발생", "E"
        GoTo RFC_Sender_Exit
    Else
        If (oFunction01.Imports("E_MESSAGE").VALUE = "") Then '//에러메시지
            ReturnValue = "Y"
            oDS_PS_MM007L.setValue "U_TransYN", i, Trim(oFunction01.Imports("E_TransYN").VALUE) '//전송접수여부
        Else
            oDS_PS_MM007L.setValue "U_MESSAGE", i, "" '초기화

            oDS_PS_MM007L.setValue "U_MESSAGE", i, Trim(oFunction01.Imports("E_MESSAGE").VALUE) '//에러메세지
            oDS_PS_MM007L.setValue "U_TransYN", i, Trim(oFunction01.Imports("E_TransYN").VALUE) '//전송접수여부
            Call MDC_Com.MDC_GF_Message(oFunction01.Imports("E_MESSAGE").VALUE, "E")
            GoTo RFC_Sender_Exit
        End If
    End If

    If Not (oSapConnection01.Connection Is Nothing) Then
        If i = LastRow Then
            oSapConnection01.Connection.Logoff
            Set oSapConnection01 = Nothing
        End If
    End If
    
    PS_RFC_Sender = ReturnValue
    Set oFunction01 = Nothing
    Exit Function
RFC_Sender_Exit:
    If Not (oSapConnection01.Connection Is Nothing) Then
        If i = LastRow Then
            oSapConnection01.Connection.Logoff
            Set oSapConnection01 = Nothing
        End If
    End If
    PS_RFC_Sender = ""
    oSeq = 99
    Set oFunction01 = Nothing
    Exit Function
RFC_Sender_Error:
    If Not (oSapConnection01.Connection Is Nothing) Then
        If i = LastRow Then
            oSapConnection01.Connection.Logoff
            Set oSapConnection01 = Nothing
        End If
    End If
    PS_RFC_Sender = ""
    oSeq = 99
    Set oFunction01 = Nothing
    Sbo_Application.SetStatusBarMessage "RFC_Sender_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
