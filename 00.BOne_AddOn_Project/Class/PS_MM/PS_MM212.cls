VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM212"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM212.cls
'//  Module         : MM
'//  Description    : 구매진도현황
'//  FormType       : PS_MM212
'//  Create Date    : 2010.12.01
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM212.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
        
    oFormUniqueID01 = "PS_MM212_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    
    oForm01.Update
           
    oForm01.Freeze False
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadForm_Error:
    oForm01.Update
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
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
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "ItemCode" Or pval.ItemUID = "CntcCode" Or pval.ItemUID = "CardCode" Then
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                
                    Print_Report01
                End If
'et_COMBO_SELECT ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                
                '청구자별로 조회할 경우에만 사번조회조건 반영(2012.07.18 송명규)
                If pval.ItemUID = "Gbn01" Then
                    
                    If oForm01.Items("Gbn01").Specific.VALUE = "10" Then
                    
                        '//아이디별 사번 세팅
                        oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
                    
                    Else
                    
                        oForm01.Items("CntcCode").Specific.VALUE = ""
                        oForm01.Items("CntcName").Specific.VALUE = ""
                        
                    End If
                
                End If
                
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    oForm01.Freeze True
                    If pval.ItemUID = "ItemCode" Or pval.ItemUID = "CntcCode" Or pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    oForm01.Freeze False
                End If
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
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
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    oForm01.Freeze False
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
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 8)
    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYY-MM") & "-01"
    
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 8)
    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    oForm01.DataSources.UserDataSources.Item("DocDateTo").VALUE = Date

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
    sQry = "SELECT U_Minor, U_CdName  From [@PS_SY001L] WHERE Code = 'C105' AND U_UseYN Like '%' ORDER BY U_Seq"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '//구분
    Set oCombo = oForm01.Items("Gbn01").Specific
    oCombo.ValidValues.Add "10", "청구기준(청구자별)"
    oCombo.ValidValues.Add "20", "품의기준(거래처별)"
    oCombo.ValidValues.Add "30", "청구기준(작번별)"
    oCombo.ValidValues.Add "40", "청구기준(분류별)"
    oCombo.Select "10", psk_ByValue
    
    '//구매방식
    Set oCombo = oForm01.Items("Purchase").Specific
    sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code"
    oRecordSet01.DoQuery sQry
    oCombo.ValidValues.Add "", ""
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//품목대분류
    Set oCombo = oForm01.Items("ItmBSort").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y' Order by Code"
    oRecordSet01.DoQuery sQry
    oCombo.ValidValues.Add "", ""
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
     '//사용처
    Set oCombo = oForm01.Items("UseDept").Specific
    sQry = "Select PrcCode, PrcName From [OPRC] Where DimCode = '1' Order by PrcCode"
    oRecordSet01.DoQuery sQry
    oCombo.ValidValues.Add "", ""
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
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

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sRow As Long
    Dim sSeq$
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
        Case "ItemCode"
            sQry = "Select FrgnName, U_Size From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            oForm01.Items("Size").Specific.VALUE = Trim(oRecordSet01.Fields(1).VALUE)
            
            
        Case "CntcCode"
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "구매견적문서가 취소되었거나 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String

    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("DocDateFr").Specific.VALUE) = "" Or Trim(oForm01.Items("DocDateTo").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    If Trim(oForm01.Items("Gbn01").Specific.VALUE) = "30" And Trim(oForm01.Items("OrdNum").Specific.VALUE) = "" Then
        ErrNum = 2
        GoTo HeaderSpaceLineDel_Error
    End If

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "품의일자는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "작번별 구매진도현황은 메인작번을 필수로 입력해야 합니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub Print_Report01()
On Error GoTo Print_Report01_Error
    Dim i&
    Dim ErrNum&
    Dim DocNum                  As String
    Dim WinTitle                As String
    Dim ReportName              As String
    Dim oText(3)                As String
    
    Dim sQry                    As String
    Dim oRecordSet01            As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId$, Gbn01$, Purchase$, ItmBsort$, ItemCode$, DocDateFr$, DocDateTo$, CntcCode$, CardCode$, OrdNum$, OrdSub1$, OrdSub2$, UseDept$
    Dim ItemName As String
    Dim Size As String
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    Gbn01 = Trim(oForm01.Items("Gbn01").Specific.VALUE)
    Purchase = Trim(oForm01.Items("Purchase").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
    OrdSub1 = Trim(oForm01.Items("OrdSub1").Specific.VALUE)
    OrdSub2 = Trim(oForm01.Items("OrdSub2").Specific.VALUE)
    UseDept = Trim(oForm01.Items("UseDept").Specific.VALUE)
    ItemName = Replace(Trim(oForm01.Items("ItemName").Specific.VALUE), "'", "''")
    Size = Replace(Trim(oForm01.Items("Size").Specific.VALUE), "'", "''") '품목명에 '가 있을 경우 처리 (2013.04.18 송명규)
    
    If BPLId = "" Then BPLId = "%"
    If Purchase = "" Then Purchase = "%"
    If ItmBsort = "" Then ItmBsort = "%"
    If ItemCode = "" Then ItemCode = "%"
    If DocDateFr = "" Then DocDateFr = "18990101"
    If DocDateTo = "" Then DocDateTo = "20991231"
    If CntcCode = "" Then CntcCode = "%"
    If CardCode = "" Then CardCode = "%"
    If OrdNum = "" Then OrdNum = "%"
    If OrdSub1 = "" Then OrdSub1 = "%"
    If OrdSub2 = "" Then OrdSub2 = "%"
    If UseDept = "" Then UseDept = "%"
    
    If Gbn01 = "10" Then
        WinTitle = "구매 진도 현황 [PS_MM212_01]"
        ReportName = "PS_MM212_01.rpt"
    ElseIf Gbn01 = "20" Then
        WinTitle = "구매 진도 현황 [PS_MM212_02]"
        ReportName = "PS_MM212_02.rpt"
    ElseIf Gbn01 = "30" Then
        WinTitle = "구매 진도 현황 [PS_MM212_03]"
        ReportName = "PS_MM212_03.rpt"
    ElseIf Gbn01 = "40" Then
        WinTitle = "구매 진도 현황 [PS_MM212_04]"
        ReportName = "PS_MM212_04.rpt"
    End If
       
    '//Formula 수식필드
    ReDim gRpt_Formula(4)
    ReDim gRpt_Formula_Value(4)
    
    oText(1) = Format(DocDateFr, "0000-00-00")
    oText(2) = Format(DocDateTo, "0000-00-00")
    sQry = "SELECT Name From [@PSH_ORDTYP] Where Code = '" & Purchase & "'"
    oRecordSet01.DoQuery sQry
    If Trim(oRecordSet01.Fields(0).VALUE) = "" Then
        oText(3) = "전체"
    Else
        oText(3) = Trim(oRecordSet01.Fields(0).VALUE)
    End If
    
    For i = 1 To 3
        If Len("" & i & "") = 1 Then
            gRpt_Formula(i) = "F0" & i & ""
        Else
            gRpt_Formula(i) = "F" & i & ""
        End If
        gRpt_Formula_Value(i) = oText(i)
    Next i
    
    
    gRpt_Formula(4) = "BPLId"
    sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'C105' AND U_Minor = '" & BPLId & "'"
    Call oRecordSet01.DoQuery(sQry)
    gRpt_Formula_Value(4) = oRecordSet01.Fields(0).VALUE
    
    '//SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '//조회조건문
    sQry = "EXEC [PS_MM212_01] '" & BPLId & "', '" & Gbn01 & "', '" & Purchase & "', '" & ItmBsort & "', '" & ItemCode & "', '" & DocDateFr & "', '" & DocDateTo & "', '" & CntcCode & "', '" & CardCode & "', '" & OrdNum & "', '" & OrdSub1 & "', '" & OrdSub2 & "', '" & UseDept & "','" & ItemName & "','" & Size & "'"
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 2
'        GoTo Print_Report01_Error
'    End If
    
    '//CR Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Or ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
