VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP050.cls
'//  Module         : PP
'//  Description    : 작업공정 내역조회(멀티)
'//  FormType       : PS_PP050
'//  Create Date    : 2010.11.30
'//  Modified Date  :
'//  Creator        : Lee Byong Gak
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP050.srf")
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
    
    oFormUniqueID01 = "PS_PP050_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    
    oForm01.Update
'
'    oForm01.Freeze False
'    oForm01.Visible = True
'    Set oXmlDoc01 = Nothing
    
    
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
           
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
'    Set oDS_PS_PP050H = oForm01.DataSources.DBDataSources("@PS_PP050H")
'    Set oDS_PS_PP050L = oForm01.DataSources.DBDataSources("@PS_PP050L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
'    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
'    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDate"
'    oForm01.DataSources.UserDataSources.Item("DocDate").Value = Format(Now, "YYYYMMDD")
    
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
       
    Call oForm01.Items("Gubun").Specific.ValidValues.Add("1", "작업이력")
    Call oForm01.Items("Gubun").Specific.ValidValues.Add("2", "검사성적서")
    oForm01.Items("Gubun").Specific.Select "0", psk_Index
              
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
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
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
                
                ElseIf pval.ItemUID = "Btn01" Then '출력버튼 클릭시
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call Print_Query
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
'                    If pval.ItemUID = "SItemCode" Then
'                        If oForm01.Items("SItemCode").Specific.Value = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
'                    If pval.ItemUID = "EItemCode" Then
'                        If oForm01.Items("EItemCode").Specific.Value = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
                    '//라인
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
            Case et_VALIDATE: '//10                                  '질의 창 띄워서 명칭 넣어주기
'                If pval.ItemChanged = True Then
'                    If pval.ItemUID = "CardCode" Then
'                       FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
'                    End If
'                End If
'
'                 If pval.ItemChanged = True Then
'                    If pval.ItemUID = "ItemMsort" Then
'                       FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
'                    End If
'                End If
'
'                  If pval.ItemChanged = True Then
'                    If pval.ItemUID = "Mark" Then
'                       FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
'                    End If
'                End If
'
'                If pval.ItemChanged = True Then
'                    If pval.ItemUID = "ItemType" Then
'                       FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
'                    End If
'                End If
            
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
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
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

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
    
    Dim CardCode$, CardName$, ItemMsort$, MsortName$, Mark$, MarkName$, ItemType$, TypeName$
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
'    Case "CardCode"
'            sQry = "SELECT CardName FROM OCRD WHERE CardCode = '" & Trim(oForm01.Items("CardCode").Specific.Value) & "'"
'            oRecordset.DoQuery sQry
'
'            oForm01.Items("CardName").Specific.String = Trim(oRecordset.Fields("CardName").Value)
'    Case "ItemMsort"
'            sQry = "SELECT U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" & Trim(oForm01.Items("ItemMsort").Specific.Value) & "'"
'            oRecordset.DoQuery sQry
'
'            oForm01.Items("MsortName").Specific.String = Trim(oRecordset.Fields("U_CodeName").Value)
'     Case "Mark"
'            sQry = "SELECT Name FROM [@PSH_MARK] WHERE Code = '" & Trim(oForm01.Items("Mark").Specific.Value) & "'"
'            oRecordset.DoQuery sQry
'
'            oForm01.Items("MarkName").Specific.String = Trim(oRecordset.Fields("Name").Value)
'     Case "ItemType"
'            sQry = "SELECT Name FROM [@PSH_SHAPE] WHERE Code= '" & Trim(oForm01.Items("ItemType").Specific.Value) & "'"
'            oRecordset.DoQuery sQry
'
'            oForm01.Items("TypeName").Specific.String = Trim(oRecordset.Fields("Name").Value)
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
'    If oCol = "Mat01" Then
'        Select Case oCol
'        Case "CardCode"
'            sQry = "SELECT CardName FROM OCRD WHERE CardCode = '" & Trim(oForm01.Items("CardName").Specific.Value) & "'"
'            oRecordSet01.DoQuery sQry
'
'            oForm01.Items("ItemNam1").Specific.String = Trim(oRecordSet01.Fields("FrgnName").Value)
'
'        Case "ItemCod2"
'            sQry = "Select FrgnName  From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCod2").Specific.Value) & "'"
'            oRecordSet01.DoQuery sQry
'
'            oForm01.Items("ItemNam2").Specific.String = Trim(oRecordSet01.Fields("FrgnName").Value)
'
'        End Select
'    End If
     
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
'        Case Trim(oDS_PS_PP050H.GetValue("U_BPLId", 0)) = ""
'            ErrNum = 1
'            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하여 주십시오.", "E"
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
    
    Dim OrdNum       As String
    Dim BatchNum       As String
    Dim Gubun       As String


    Dim PackingNo       As String
    Dim ItemCode        As String
    Dim CardCode        As String
    Dim Chk             As String
    Dim BakeTestChk As String 'BakeTest 측정치 추가 품목 체크를 위한 변수(2012.11.26 송명규 추가)
    Dim PackNo As String '패킹번호(2012.11.26 송명규 추가)
    
    
    
    Dim BPLID$
   
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// 조회조건문
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
    BatchNum = Trim(oForm01.Items("BatchNum").Specific.VALUE)
    Gubun = Trim(oForm01.Items("Gubun").Specific.VALUE)
    
    If OrdNum = "" And BatchNum = "" Then
        MDC_Com.MDC_GF_Message "관리번호 및 온산 LotNo 하나 이상은 입력해야합니다.", "E"
        ErrNum = 1
        GoTo Print_Query_Error
    End If
    
    If OrdNum = "" Then OrdNum = "%"
    If BatchNum = "" Then BatchNum = "%"
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", oRecordSet.RecordCount, False)
            '/ Action (sub_query가 있을때는 'Y'로...)/
            
            
    If Gubun = 1 Then
        '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        WinTitle = "[PS_PP050_01] 작업공정 내역조회(멀티)"
        ReportName = "PS_PP050_01.RPT"
        
        '// Formula 수식필드
        ReDim gRpt_Formula(2)
        ReDim gRpt_Formula_Value(2)
        
    '    gRpt_Formula(1) = "DocDateFr"
    '    gRpt_Formula_Value(1) = DocDate
        
        '// SubReport
        ReDim gRpt_SRptSqry(1)
        ReDim gRpt_SRptName(1)
        
        ReDim gRpt_SFormula(1, 1)
        ReDim gRpt_SFormula_Value(1, 1)
        
        gRpt_SFormula(1, 1) = ""
        gRpt_SFormula_Value(1, 1) = ""
        
        '/ Procedure 실행"
        sQry = "EXEC [PS_PP050_01] '" & BPLID & "', '" & OrdNum & "', '" & BatchNum & "'"
                 
        
        If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        End If
    
    
     Else    ' 검사성적서 출력 구문
        
        sQry = "Exec [PS_QM041_06] '" & BPLID & "', '" & OrdNum & "', '" & BatchNum & "'"  '삭제 로직 포함

        oRecordSet.DoQuery sQry
      '  oMat01.FlushToDataSource
      
        sQry = "Exec [PS_PP050_02] '" & BPLID & "', '" & OrdNum & "', '" & BatchNum & "'"  '
        
        oRecordSet.DoQuery sQry
        
        ItemCode = MDC_PS_Common.GetValue("Select b.U_ItemCode From [@PS_PP090H] a inner Join [@PS_PP090L] b on a.DocEntry = b.DocEntry left join [@PS_QM020H] c on b.U_LotNo = c.U_OrdNum left join  obtn e on b.u_lotno = e.DistNumber Where a.U_BPLId = '" & BPLID & "' and e.u_packno like '" & OrdNum & "' + '%' and U_BaseLot like '" & BatchNum & "' + '%' AND a.Canceled = 'N' Group by b.U_ItemCode, c.U_CardCode ")
        CardCode = MDC_PS_Common.GetValue("Select c.U_CardCode From [@PS_PP090H] a inner Join [@PS_PP090L] b on a.DocEntry = b.DocEntry left join [@PS_QM020H] c on b.U_LotNo = c.U_OrdNum left join  obtn e on b.u_lotno = e.DistNumber Where a.U_BPLId = '" & BPLID & "' and e.u_packno like '" & OrdNum & "' + '%' and U_BaseLot like '" & BatchNum & "' + '%' AND a.Canceled = 'N' Group by b.U_ItemCode, c.U_CardCode ")
        
        
    WinTitle = "[PS_QM041] 검사성적서출력(신양식)"
                                                                                                                                                                                                                                                
    If Chk = "Y" Then
       ReportName = "PS_QM041_02.RPT"  '// B/G 타입
    Else
    
        If CardCode <> "12440" Then
            ReportName = "PS_QM041_01.RPT"  '// 일반
        Else
        
            If ItemCode <> "104010098" Then '응급조치로 ItemCode로 체크(추후 품목이 추가될 경우를 생각해서 마스터데이터의 특정 값 기준으로 수정 할 필요 있음, 2012.11.26 송명규)
                ReportName = "PS_QM041_03.RPT"  '// (주)TSD
            Else
                ReportName = "PS_QM041_04.RPT"  '// (주)TSD
            End If
          
       End If
       
    End If
                                                                                                                                                                                                                                                            
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '// SubReport
    ReDim gRpt_SRptSqry(3)
    ReDim gRpt_SRptName(3)
    
    ReDim gRpt_SFormula(3, 1)
    ReDim gRpt_SFormula_Value(3, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    gRpt_SFormula(2, 1) = ""
    gRpt_SFormula_Value(2, 1) = ""
    
    gRpt_SFormula(3, 1) = ""
    gRpt_SFormula_Value(3, 1) = ""
        
    If Chk = "Y" Then
           
       Sub_sQry = "EXEC [PS_QM041_04] '" & BPLID & "'"
       gRpt_SRptSqry(1) = Sub_sQry
       gRpt_SRptName(1) = "PS_QM040_SUB_01"
       
       Sub_sQry = "EXEC [PS_QM041_05] '" & BPLID & "'"
       gRpt_SRptSqry(2) = Sub_sQry
       gRpt_SRptName(2) = "PS_QM040_SUB_02"
       
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLID & "'"
       gRpt_SRptSqry(3) = Sub_sQry
       gRpt_SRptName(3) = "PS_QM040_SUB_03"
       
    Else
    
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLID & "'"
       gRpt_SRptSqry(1) = Sub_sQry
       gRpt_SRptName(1) = "PS_QM040_SUB_01"
       
       Sub_sQry = "EXEC [PS_QM041_05] '" & BPLID & "'"
       gRpt_SRptSqry(2) = Sub_sQry
       gRpt_SRptName(2) = "PS_QM040_SUB_02"
       
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLID & "'"
       gRpt_SRptSqry(3) = Sub_sQry
       gRpt_SRptName(3) = "PS_QM040_SUB_03"
       
    End If

    
        sQry = "EXEC [PS_QM041_02] '" & BPLID & "'"
        Call oRecordSet.DoQuery(sQry)
            
             '/ Action (sub_query가 있을때는 'Y'로...)/
        If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "", "N", "V") = False Then
        End If
    
     End If



    
    ProgBar01.VALUE = 100
'    ProgBar01.Text = "비용 집계중...!"
    ProgBar01.Stop
    
    Set ProgBar01 = Nothing
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set ProgBar01 = Nothing
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
