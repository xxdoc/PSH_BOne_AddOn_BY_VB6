VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP400.cls
'//  Module         : PP
'//  Description    : 납기초과 생산품 현황
'//  FormType       : PS_PP400
'//  Create Date    : 2011.02.16
'//  Modified Date  :
'//  Creator        : Lee Byong Gak
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oGrid01 As SAPbouiCOM.Grid
Public oGrid02 As SAPbouiCOM.Grid

Public oDS_PS_PP400A As SAPbouiCOM.DataTable
Public oDS_PS_PP400B As SAPbouiCOM.DataTable

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP400.srf")
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
    
    oFormUniqueID01 = "PS_PP400_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), False        '// 행삭제

    oForm01.Items("FrDt01").Specific.VALUE = Format(Now, "YYYY0101")
    oForm01.Items("ToDt01").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oForm01.Items("FrDt02").Specific.VALUE = Format(Now, "YYYYMM01")
    oForm01.Items("ToDt02").Specific.VALUE = Format(Now, "YYYYMMDD")

    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨

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

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//아이디별 사업장 세팅
    oForm01.Items("BPLID01").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
    oForm01.Items("BPLID02").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
             
'Call oForm01.Items("Pumtxt").Specific.ValidValues.Add("", "")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("Pumtxt").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", False, False)
             
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


Private Sub CreateItems()
On Error GoTo CreateItems_Error

    Set oGrid01 = oForm01.Items("Grid01").Specific
    Set oGrid02 = oForm01.Items("Grid02").Specific
    
    oGrid01.SelectionMode = ms_Single
    oGrid02.SelectionMode = ms_Single
    
    Call oForm01.DataSources.DataTables.Add("PS_PP400A")
    Call oForm01.DataSources.DataTables.Add("PS_PP400B")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("PS_PP400A")
    oGrid02.DataTable = oForm01.DataSources.DataTables.Item("PS_PP400B")
    
    Set oDS_PS_PP400A = oForm01.DataSources.DataTables.Item("PS_PP400A")
    Set oDS_PS_PP400B = oForm01.DataSources.DataTables.Item("PS_PP400B")

    '//////////납기도래품목 조회//////////_S
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID01", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID01").Specific.DataBind.SetBound(True, "", "BPLID01")

    '기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt01", dt_DATE)
    Call oForm01.Items("FrDt01").Specific.DataBind.SetBound(True, "", "FrDt01")
    
    '기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt01", dt_DATE)
    Call oForm01.Items("ToDt01").Specific.DataBind.SetBound(True, "", "ToDt01")
    
    '(작번)품목코드
    Call oForm01.DataSources.UserDataSources.Add("ItemCode01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemCode01").Specific.DataBind.SetBound(True, "", "ItemCode01")
    
    '품목명
    Call oForm01.DataSources.UserDataSources.Add("ItemName01", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName01").Specific.DataBind.SetBound(True, "", "ItemName01")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("ItemSpec01", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemSpec01").Specific.DataBind.SetBound(True, "", "ItemSpec01")
    
    '거래처코드
    Call oForm01.DataSources.UserDataSources.Add("CardCode01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardCode01").Specific.DataBind.SetBound(True, "", "CardCode01")
    
    '거래처명
    Call oForm01.DataSources.UserDataSources.Add("CardName01", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CardName01").Specific.DataBind.SetBound(True, "", "CardName01")
    
    '거래처구분
    Call oForm01.DataSources.UserDataSources.Add("CardType01", dt_SHORT_TEXT, 10)
    Call oForm01.Items("CardType01").Specific.DataBind.SetBound(True, "", "CardType01")
    
    '품목구분
    Call oForm01.DataSources.UserDataSources.Add("ItemType01", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItemType01").Specific.DataBind.SetBound(True, "", "ItemType01")
    '//////////납기도래품목 조회//////////_E
    
    '//////////납기초과생산품 조회//////////_S
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLID02", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID02").Specific.DataBind.SetBound(True, "", "BPLID02")

    '기간(시작)
    Call oForm01.DataSources.UserDataSources.Add("FrDt02", dt_DATE)
    Call oForm01.Items("FrDt02").Specific.DataBind.SetBound(True, "", "FrDt02")
    
    '기간(종료)
    Call oForm01.DataSources.UserDataSources.Add("ToDt02", dt_DATE)
    Call oForm01.Items("ToDt02").Specific.DataBind.SetBound(True, "", "ToDt02")
    
    '(작번)품목코드
    Call oForm01.DataSources.UserDataSources.Add("ItemCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("ItemCode02").Specific.DataBind.SetBound(True, "", "ItemCode02")
    
    '품목명
    Call oForm01.DataSources.UserDataSources.Add("ItemName02", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName02").Specific.DataBind.SetBound(True, "", "ItemName02")
    
    '규격
    Call oForm01.DataSources.UserDataSources.Add("ItemSpec02", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemSpec02").Specific.DataBind.SetBound(True, "", "ItemSpec02")
    
    '거래처코드
    Call oForm01.DataSources.UserDataSources.Add("CardCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CardCode02").Specific.DataBind.SetBound(True, "", "CardCode02")
    
    '거래처명
    Call oForm01.DataSources.UserDataSources.Add("CardName02", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CardName02").Specific.DataBind.SetBound(True, "", "CardName02")
    '//////////납기초과생산품 조회//////////_E
    
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

    '//////////납기도래품목 조회//////////_S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID01").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, False, False)
    
    '거래처구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'C100'"
    Call oForm01.Items("CardType01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType01").Specific, sQry, "%", False, False)
    
    '품목구분
    sQry = "           SELECT     U_Minor AS [Code], "
    sQry = sQry & "               U_CdName AS [Name]"
    sQry = sQry & " FROM      [@PS_SY001L]"
    sQry = sQry & " WHERE     Code = 'S002'"
    Call oForm01.Items("ItemType01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemType01").Specific, sQry, "%", False, False)
    '//////////납기도래품목 조회//////////_E
    
    '//////////납기초과생산품 조회//////////_S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID02").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", User_BPLId, False, False)
    '//////////납기초과생산품 조회//////////_E

'    Set oCombo = oForm01.Items("CardCode").Specific
'    Call oCombo.ValidValues.Add("0", "전체선택")
      
'    Set oCombo = oForm01.Items("Pumtxt").Specific
'    Call oCombo.ValidValues.Add("0", "전체선택")

'    Set oCombo = oForm01.Items("ItMBsort").Specific
'    sQry = "SELECT Code, Name From [@PSH_ItmBsort] Where Code in ('101', '102', '107')"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
       
'    Call oForm01.Freeze(True)
    
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
                    
                ElseIf pval.ItemUID = "BtnSrch01" Then '납기초과생산품 조회
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call PS_PP400_MTX01
                    End If
                
                ElseIf pval.ItemUID = "BtnSrch02" Then '납기초과생산품 조회
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call PS_PP400_MTX02
                    End If
                   
                ElseIf pval.ItemUID = "BtnPrt01" Then '납기초과생산품 조회(리포트)
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call Print_Query01
                    End If

                ElseIf pval.ItemUID = "BtnPrt02" Then '납기초과생산품 조회(리포트)
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call Print_Query02
                    End If
                End If
            Case et_KEY_DOWN: '//2
            
                If pval.CharPressed = 9 Then
                
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode01", "")
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode01", "")
                    
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode02", "")
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode02", "")
                
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



                '폴더를 사용할 때는 필수 소스_S
                'Folder01이 선택되었을 때
                If pval.ItemUID = "Folder01" Then
                    
                    oForm01.PaneLevel = 1
                    oForm01.DefButton = "BtnSrch01"
                    
                End If
                
                'Folder02가 선택되었을 때
                If pval.ItemUID = "Folder02" Then
                    
                    oForm01.PaneLevel = 2
                    oForm01.DefButton = "BtnSrch02"
                    
                End If
                '폴더를 사용할 때는 필수 소스_E

            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            
                If pval.ItemChanged = True Then
                    Call FlushToItemValue(pval.ItemUID)
                End If
            
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            
                '그룹박스 크기 동적 할당
                oForm01.Items("GrpBox01").Height = oForm01.Items("Grid01").Height + 110
                oForm01.Items("GrpBox01").Width = oForm01.Items("Grid01").Width + 30
                
                If oGrid01.Columns.Count > 0 Then
                    Call oGrid01.AutoResizeColumns
                End If
                
                If oGrid02.Columns.Count > 0 Then
                    oGrid02.AutoResizeColumns
                End If
            
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
    
    Select Case oUID
    
        Case "ItemCode01"
        
            oForm01.Items("ItemName01").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "OITM", "'" & Trim(oForm01.Items("ItemCode01").Specific.VALUE) & "'") '작번
            oForm01.Items("ItemSpec01").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "OITM", "'" & Trim(oForm01.Items("ItemCode01").Specific.VALUE) & "'") '규격
        
        Case "CardCode01"
        
            oForm01.Items("CardName01").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "OCRD", "'" & Trim(oForm01.Items("CardCode01").Specific.VALUE) & "'") '거래처
    
        Case "ItemCode02"
        
            oForm01.Items("ItemName02").Specific.VALUE = MDC_GetData.Get_ReData("FrgnName", "ItemCode", "OITM", "'" & Trim(oForm01.Items("ItemCode02").Specific.VALUE) & "'") '작번
            oForm01.Items("ItemSpec02").Specific.VALUE = MDC_GetData.Get_ReData("U_Size", "ItemCode", "OITM", "'" & Trim(oForm01.Items("ItemCode02").Specific.VALUE) & "'") '규격
        
        Case "CardCode02"
        
            oForm01.Items("CardName02").Specific.VALUE = MDC_GetData.Get_ReData("CardName", "CardCode", "OCRD", "'" & Trim(oForm01.Items("CardCode02").Specific.VALUE) & "'") '거래처
    
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
'        Case Trim(oDS_PS_PP400H.GetValue("U_BPLId", 0)) = ""
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

Private Sub PS_PP400_MTX01()
On Error GoTo PS_PP400_MTX01_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim Query02 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String '사업장
    Dim CardCode As String
    Dim ItemCode As String
    Dim FrDt As String '기간(Fr)
    Dim ToDt As String '기간(To)
    Dim CardType As String '거래처구분
    Dim ItemType As String '품목구분
    
    BPLID = Trim(oForm01.Items("BPLID01").Specific.Selected.VALUE) '사업장
    CardCode = Trim(oForm01.Items("CardCode01").Specific.VALUE) '거래처
    ItemCode = Trim(oForm01.Items("ItemCode01").Specific.VALUE) '품목코드
    FrDt = Trim(oForm01.Items("FrDt01").Specific.VALUE) '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt01").Specific.VALUE) '기간(To)
    CardType = Trim(oForm01.Items("CardType01").Specific.Selected.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("ItemType01").Specific.Selected.VALUE) '품목구분
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
   
    Query01 = "              EXEC PS_PP400_01 '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & CardCode & "','"
    Query01 = Query01 & ItemCode & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "','"
    Query01 = Query01 & CardType & "','"
    Query01 = Query01 & ItemType & "'"
    
    Call oGrid01.DataTable.Clear
    Call oDS_PS_PP400A.ExecuteQuery(Query01)
'    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid01.Columns(7).RightJustified = True
    oGrid01.Columns(8).RightJustified = True
    oGrid01.Columns(9).RightJustified = True
    oGrid01.Columns(13).RightJustified = True
    oGrid01.Columns(18).RightJustified = True
    oGrid01.Columns(20).RightJustified = True
    oGrid01.Columns(21).RightJustified = True
    
    oGrid01.Columns(13).BackColor = RGB(255, 255, 125) '잔여납기일, 노랑
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '[결산]계, 노랑
'    oGrid01.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid01.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid01.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid01.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid01.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid01.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid01.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid01.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid01.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP400_MTX01_Exit
    End If

    oGrid01.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP400_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP400_MTX01_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_PP400_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP400_MTX02()
'******************************************************************************
'Function ID : PS_PP400_MTX02()
'해당모듈 : PS_PP400
'기능 : 납기초과생산품 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP400_MTX02_Error

    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim Query02 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String '사업장
    Dim FrDt As String '기간(Fr)
    Dim ToDt As String '기간(To)
    Dim ItemCode As String '공정
    Dim CardCode As String '작업구분
    
    BPLID = Trim(oForm01.Items("BPLID02").Specific.Selected.VALUE) '사업장
    FrDt = Trim(oForm01.Items("FrDt02").Specific.VALUE) '기간(Fr)
    ToDt = Trim(oForm01.Items("ToDt02").Specific.VALUE) '기간(To)
    ItemCode = Trim(oForm01.Items("ItemCode02").Specific.VALUE) '작번(품목코드)
    CardCode = Trim(oForm01.Items("CardCode02").Specific.VALUE) '거래처코드
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
   
    Query01 = "              EXEC [PS_PP400_02] '"
    Query01 = Query01 & BPLID & "','"
    Query01 = Query01 & CardCode & "','"
    Query01 = Query01 & ItemCode & "','"
    Query01 = Query01 & FrDt & "','"
    Query01 = Query01 & ToDt & "'"
    
    Call oGrid02.DataTable.Clear
    Call oDS_PS_PP400B.ExecuteQuery(Query01)
'    oGrid02.DataTable = oForm01.DataSources.DataTables.Item("DataTable")

    oGrid02.Columns(8).RightJustified = True
    oGrid02.Columns(9).RightJustified = True
    oGrid02.Columns(10).RightJustified = True
    oGrid02.Columns(11).RightJustified = True
    oGrid02.Columns(14).RightJustified = True
'    oGrid02.Columns(14).RightJustified = True
'    oGrid02.Columns(15).RightJustified = True
    
    
'    oGrid02.Columns(19).BackColor = RGB(255, 255, 125) '[계산]계, 노랑
'    oGrid02.Columns(26).BackColor = RGB(255, 255, 125) '[완료]계, 노랑
    
'    oGrid02.Columns(9).BackColor = RGB(255, 255, 125) '품의일, 노랑
'    oGrid02.Columns(10).BackColor = RGB(255, 255, 125) '가입고일, 노랑
'    oGrid02.Columns(11).BackColor = RGB(0, 210, 255) '차이(품의-가입고), 하늘
'    oGrid02.Columns(12).BackColor = RGB(255, 255, 125) '검수입고일, 노랑
'    oGrid02.Columns(13).BackColor = RGB(0, 210, 255) '차이(가입고-품의), 하늘
'    oGrid02.Columns(14).BackColor = RGB(255, 167, 167) '총소요일, 빨강
        
    If oGrid02.Rows.Count = 0 Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_PP400_MTX02_Exit
    End If

    oGrid02.AutoResizeColumns
    oForm01.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP400_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP400_MTX02_Error:

    Call oForm01.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    Sbo_Application.SetStatusBarMessage "PS_PP400_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Print_Query01()
'******************************************************************************
'Function ID : Print_Query01()
'해당모듈 : PS_PP400
'기능 : 납기도래품목 조회(리포트)
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo Print_Query01_Error

    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry        As String
    
    Dim BPLID As String
    Dim CardCode As String
    Dim ItemCode As String
    Dim FrDt As String
    Dim ToDt As String
    Dim CardType As String
    Dim ItemType As String
   
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
   
    Call ConnectODBC
    
    '// 조회조건문
    BPLID = Trim(oForm01.Items("BPLID01").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode01").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode01").Specific.VALUE)
    FrDt = Trim(oForm01.Items("FrDt01").Specific.VALUE)
    ToDt = Trim(oForm01.Items("ToDt01").Specific.VALUE)
    CardType = Trim(oForm01.Items("CardType01").Specific.Selected.VALUE) '거래처구분
    ItemType = Trim(oForm01.Items("ItemType01").Specific.Selected.VALUE) '품목구분
        
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP400_01] 납기도래품목 조회"

    ReportName = "PS_PP400_01.RPT"

    '// Formula 수식필드
    ReDim gRpt_Formula(0)
    ReDim gRpt_Formula_Value(0)
    
'    gRpt_Formula(1) = "DocDateFr"
'    gRpt_Formula_Value(1) = Left(DocDateFr, 4) + "-" + Mid(DocDateFr, 5, 2) + "-" + Right(DocDateFr, 2)
'    gRpt_Formula(2) = "DocDateTo"
'    gRpt_Formula_Value(2) = Left(DocDateTo, 4) + "-" + Mid(DocDateTo, 5, 2) + "-" + Right(DocDateTo, 2)
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
'    '/ Procedure 실행"

    sQry = "         EXEC PS_PP400_51 '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & CardCode & "','"
    sQry = sQry & ItemCode & "','"
    sQry = sQry & FrDt & "','"
    sQry = sQry & ToDt & "','"
    sQry = sQry & CardType & "','"
    sQry = sQry & ItemType & "'"
       
'    oRecordset.DoQuery sQry
'    If oRecordset.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub


Private Sub Print_Query02()
'******************************************************************************
'Function ID : Print_Query02()
'해당모듈 : PS_PP400
'기능 : 납기초과생산품 조회(리포트)
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo Print_Query02_Error

    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry        As String
    
    Dim BPLID       As String
    Dim CardCode    As String
    Dim ItemCode    As String
    Dim DocDateFr       As String
    Dim DocDateTo       As String

    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
   
    Call ConnectODBC
    
    '// 조회조건문
    BPLID = Trim(oForm01.Items("BPLID02").Specific.VALUE)
    CardCode = Trim(oForm01.Items("CardCode02").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode02").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("FrDt02").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("ToDt02").Specific.VALUE)
        
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP400_01] 납기초과생산품 조회"

    ReportName = "PS_PP400_02.RPT"

    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
    gRpt_Formula(1) = "DocDateFr"
    gRpt_Formula_Value(1) = Left(DocDateFr, 4) + "-" + Mid(DocDateFr, 5, 2) + "-" + Right(DocDateFr, 2)
    gRpt_Formula(2) = "DocDateTo"
    gRpt_Formula_Value(2) = Left(DocDateTo, 4) + "-" + Mid(DocDateTo, 5, 2) + "-" + Right(DocDateTo, 2)
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
'    '/ Procedure 실행"

    sQry = "         EXEC [PS_PP400_52] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & CardCode & "','"
    sQry = sQry & ItemCode & "','"
    sQry = sQry & DocDateFr & "','"
    sQry = sQry & DocDateTo & "'"
       
'    oRecordset.DoQuery sQry
'    If oRecordset.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query02_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub






