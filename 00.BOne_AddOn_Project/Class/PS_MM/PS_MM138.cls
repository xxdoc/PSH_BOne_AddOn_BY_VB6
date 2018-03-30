VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM138"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM138.cls
'//  Module         : MM
'//  Description    : 포장사업팀 외주입고
'//  FormType       : PS_MM138
'//  Create Date    : 2013.04.17
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_MM138H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM138L As SAPbouiCOM.DBDataSource    '등록라인

Private Type PS_PP040s
    PP030HNo As String
    PP030MNo As String
    OrdNum As String
    OrdGbn As String
    ItemCode As String
    ItemName As String
    Sequence As String
    CpCode As String
    CpName As String
    Chk As String
End Type

Private PS_PP040() As PS_PP040s
    
Private Type PS_PP040DocEntrys
    PP040DocEntry As String
End Type

Private PS_PP040DocEntry() As PS_PP040DocEntrys

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

Private oBPLId As String
Private oDocdate As String
Private oCardCode As String
Private oCardName As String

Private oDocEntryNext As Long
Private oCheck As String

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM138.srf")
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
    
    oFormUniqueID01 = "PS_MM138_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormClear
'    Call Add_MatrixRow(0, True)
    Call SetDocument(oFromDocEntry01)
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), False        '// 행삭제
        
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
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim Calculate_Weight As Double
    
    Dim DocEntry As Long
    Dim DocEntryMax As Long
    Dim DocEntryNext As Long
    
    Dim ScrapWt As Currency
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        oBPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
                        oDocdate = Trim(oForm01.Items("DocDate").Specific.VALUE)
                        oCardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
                        oCardName = Trim(oForm01.Items("CardName").Specific.VALUE)
                    
'                        If Trim(Sbo_Company.UserName) = "66302" Or Trim(Sbo_Company.UserName) = "67517" Then
'                        Else
'                            MDC_Com.MDC_GF_Message "해당 아이디는 해당 문서를 추가할 수 없습니다.", "E"
'                            BubbleEvent = False
'                            Exit Sub
'                        End If

                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        
                        '문서만 생성할때 막음 시작
                        If PS_MM138_DI_API01 = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        '문서만 생성할때 막음 끝
'
                        Call Delete_EmptyRow
                    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        '입고문서가 생성안되었을때 DI처리
                        If Trim(oForm01.Items("OIGNNo").Specific.VALUE) = "" Then
                            If PS_MM138_DI_API01 = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                        
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        
                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "OutDoc" Then
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ItmGrpCd" Then
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ItemCode" Then
                        If oForm01.Items(pval.ItemUID).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                        
                    ElseIf pval.ItemUID = "CpCode" Then
                        If oForm01.Items("CpCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "OtDocLin" Then
                            If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE = "" Then
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
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                   
                    
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                    
                    
                        
                ElseIf pval.ItemUID = "Btn01" Then
                    Print_Report01
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "IssueYN" Then
                    If oForm01.Items("IssueYN").Specific.VALUE = "Y" Then
                        If oMat01.VisualRowCount < 1 Then
                            Add_MatrixRow 0, True
                            oMat01.Columns("ItemCode").Cells(1).Specific.VALUE = Trim(oForm01.Items("ItemCode").Specific.VALUE)
                            oMat01.Columns("ItemName").Cells(1).Specific.VALUE = Trim(oForm01.Items("ItemName").Specific.VALUE)
                        End If
                        
                    Else
                        If oForm01.Mode = fm_ADD_MODE Then
                            FlushToItemValue "ItemCode"
                        End If
                    End If
                End If
            Case et_CLICK: '//6
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "OutDoc" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ItmGrpCd" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "CpCode" Then
                        FlushToItemValue pval.ItemUID
                        
                    ElseIf pval.ItemUID = "Qty" Then
                        If oMat01.VisualRowCount > 0 Then
                            oForm01.Freeze True
                            If (oForm01.Items("BPLId").Specific.VALUE = "3" Or oForm01.Items("BPLId").Specific.VALUE = "5") And oForm01.Items("ItmGrpCd").Specific.VALUE <> "" Then
                                oMat01.FlushToDataSource
                                For i = 0 To oMat01.VisualRowCount - 1
                                    If Trim(oDS_PS_MM138L.GetValue("U_ItemCode", i)) <> "" Then
                                        Call oDS_PS_MM138L.setValue("U_Qty", i, oForm01.Items("Qty").Specific.VALUE)
            
        '                                If Left(Trim(oDS_PS_MM138H.GetValue("U_ItmGrpCd", 0)), 1) = "A" Or Trim(oDS_PS_MM138H.GetValue("U_BPLId", 0)) = "5" Then
                                            '//제품 Set기준
                                            
                                            If CLng(oDS_PS_MM138L.GetValue("U_Qty", i)) = CLng(oDS_PS_MM138L.GetValue("U_StdQty", i)) Then
                                                Call oDS_PS_MM138L.setValue("U_OutQty", i, oDS_PS_MM138L.GetValue("U_BOMQty", i))
                                                Call oDS_PS_MM138L.setValue("U_OutWt", i, oDS_PS_MM138L.GetValue("U_BOMWt", i))
                                            Else
                                                If Left(Trim(oDS_PS_MM138H.GetValue("U_ItmGrpCd", 0)), 1) = "B" Then
                                                    Call oDS_PS_MM138L.setValue("U_OutQty", i, oDS_PS_MM138L.GetValue("U_BOMQty", i) * (oDS_PS_MM138L.GetValue("U_Qty", i) / oDS_PS_MM138L.GetValue("U_StdQty", i)) / 1000)
                                                    Call oDS_PS_MM138L.setValue("U_OutWt", i, oDS_PS_MM138L.GetValue("U_BOMWt", i) * (oDS_PS_MM138L.GetValue("U_Qty", i) / oDS_PS_MM138L.GetValue("U_StdQty", i)) / 1000)
                                                Else
                                                    Call oDS_PS_MM138L.setValue("U_OutQty", i, oDS_PS_MM138L.GetValue("U_BOMQty", i) * (oDS_PS_MM138L.GetValue("U_Qty", i) / oDS_PS_MM138L.GetValue("U_StdQty", i)))
                                                    Call oDS_PS_MM138L.setValue("U_OutWt", i, oDS_PS_MM138L.GetValue("U_BOMWt", i) * (oDS_PS_MM138L.GetValue("U_Qty", i) / oDS_PS_MM138L.GetValue("U_StdQty", i)))
                                                End If
                                            End If
        '                                Else
        '                                    '//원재료 기준
        '                                    sQry = "Select U_UnWeight From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE) & "'"
        '                                    oRecordSet01.DoQuery sQry
        '
        '                                    If oRecordSet01.Fields(0).VALUE > 0 Then
        '                                         Call oDS_PS_MM138L.setValue("U_OutWt", i, oDS_PS_MM138L.GetValue("U_Qty", i) * (oRecordSet01.Fields(0).VALUE / 1000))
        '                                         Call oDS_PS_MM138L.setValue("U_OutQty", i, oDS_PS_MM138L.GetValue("U_BOMQty", i) * oDS_PS_MM138L.GetValue("U_OutWt", i) / oDS_PS_MM138L.GetValue("U_BOMWt", i))
        '                                    End If
        '                                End If
                                    End If
                                Next
                                
                                oMat01.LoadFromDataSource
                                oMat01.Columns("OutQty").Cells(1).Click ct_Regular
                            ElseIf oForm01.Items("BPLId").Specific.VALUE = "3" And oForm01.Items("ItmGrpCd").Specific.VALUE = "" Then
                                oForm01.Items("Qty").Specific.VALUE = 0
                                MDC_Com.MDC_GF_Message "외주관리그룹을 입력하시기 바랍니다.", "E"
                            End If
                            oForm01.Freeze False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                    
                        If pval.ColUID = "OtDocLin" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        End If
                        
                        
                    End If
                End If
'et_MATRIX_LOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_MATRIX_LOAD: '//11

                            
'                Add_MatrixRow oMat01.RowCount, False
                FormItemEnabled
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
                Set oDS_PS_MM138H = Nothing
                Set oDS_PS_MM138L = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set oRecordSet01 = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim DocEntry As Long
    Dim DocEntryMax As Long
    Dim DocEntryNext As Long
    
    Dim OutDoc As Long
    Dim BPLID As String
    Dim OIGNNo As String
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                If (oForm01.Mode <> fm_ADD_MODE) Then
                    OIGNNo = Trim(oForm01.Items("OIGNNo").Specific.VALUE)
                    
                    If Trim(OIGNNo) <> "" Then
                    
                        BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
                        OutDoc = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                        
                        sQry = "Select Count(*) from [@PS_MM095H] a Inner Join [@PS_MM095L] b On a.DocEntry = b.DocEntry and a.Canceled = 'N' "
                        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
                        sQry = sQry + " and b.U_WorkGbn = '30' " '외주
                        sQry = sQry + " and b.U_PP040Doc = '" & OutDoc & "'"
                        
                        oRecordSet01.DoQuery (sQry)
                        
                        If Trim(oRecordSet01.Fields(0).VALUE) <= 0 Then
                            
                            If Sbo_Application.MessageBox("정말로 취소하시겠습니까?", "1", "예", "아니오") <> "1" Then
                                    BubbleEvent = False
                                    Exit Sub
                            End If
                            
                            
                            If PS_MM138_DI_API02 = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            Sbo_Application.MessageBox ("원재료 불출이 되어 취소할 수 없습니다.")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                Else
                    Sbo_Application.MessageBox ("신규모드일때는 취소할 수 없습니다..")
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                If Trim(oForm01.Items("DocEntry").Specific.VALUE) = "" Then
                    DocEntry = 0
                Else
                    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                End If
                sQry = "Select Max(DocEntry) From [@PS_MM138H]"
                oRecordSet01.DoQuery sQry
                DocEntryMax = Trim(oRecordSet01.Fields(0).VALUE)
                
                If pval.MenuUID = "1288" Then '//다음
                    If oForm01.Mode = fm_ADD_MODE Then
                        Sbo_Application.ActivateMenuItem ("1290")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        If (oForm01.Items("DocEntry").Specific.VALUE = "") Then
                            Sbo_Application.ActivateMenuItem ("1290")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
One_More_Check_1288:
                    DocEntryNext = DocEntry + 1
                    If DocEntryNext > DocEntryMax Then
                        DocEntry = 0
                        GoTo One_More_Check_1288
                    End If
                        
                    sQry = "Select U_CardCode From [@PS_MM138H] Where DocEntry = '" & DocEntryNext & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                        If Left(Trim(Sbo_Company.UserName), 1) = "6" Then
                            DocEntry = DocEntry + 1
                            GoTo One_More_Check_1288
                        Else
                            oCheck = "False"
                            Exit Sub
                        End If
                    Else
                        oCheck = "True"
                        oDocEntryNext = DocEntryNext
                    End If
                ElseIf pval.MenuUID = "1289" Then '//이전
                    If oForm01.Mode = fm_ADD_MODE Then
                        Sbo_Application.ActivateMenuItem ("1291")
                        BubbleEvent = False
                        Exit Sub
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        If (oForm01.Items("DocEntry").Specific.VALUE = "") Then
                            Sbo_Application.ActivateMenuItem ("1291")
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
One_More_Check_1289:
                    DocEntryNext = DocEntry - 1
                    If DocEntryNext < 1 Then
                        DocEntry = DocEntryMax + 1
                        GoTo One_More_Check_1289
                    End If
                        
                    sQry = "Select U_CardCode From [@PS_MM138H] Where DocEntry = '" & DocEntryNext & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                        If Left(Trim(Sbo_Company.UserName), 1) = "6" Then
                            DocEntry = DocEntry - 1
                            GoTo One_More_Check_1289
                        Else
                            oCheck = "False"
                            Exit Sub
                        End If
                    Else
                        oCheck = "True"
                        oDocEntryNext = DocEntryNext
                    End If
                ElseIf pval.MenuUID = "1290" Then '//맨첨
                    DocEntryNext = 0
One_More_Check_1290:
                    DocEntryNext = DocEntryNext + 1
                    
                    sQry = "Select U_CardCode From [@PS_MM138H] Where DocEntry = '" & DocEntryNext & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                        If Left(Trim(Sbo_Company.UserName), 1) = "6" Then
                            GoTo One_More_Check_1290
                        Else
                            oCheck = "False"
                            Exit Sub
                        End If
                    Else
                        oCheck = "True"
                        oDocEntryNext = DocEntryNext
                    End If
                ElseIf pval.MenuUID = "1291" Then '//맨뒤
                    DocEntryNext = DocEntryMax + 1
One_More_Check_1291:
                    DocEntryNext = DocEntryNext - 1
                    
                    sQry = "Select U_CardCode From [@PS_MM138H] Where DocEntry = '" & DocEntryNext & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If Trim(Sbo_Company.UserName) <> Trim(oRecordSet01.Fields(0).VALUE) Then
                        If Left(Trim(Sbo_Company.UserName), 1) = "6" Then
                            GoTo One_More_Check_1291
                        Else
                            oCheck = "False"
                            Exit Sub
                        End If
                    Else
                        oCheck = "True"
                        oDocEntryNext = DocEntryNext
                    End If
                End If
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제

            Case "1281": '찾기
                oForm01.Freeze True
                oDS_PS_MM138H.setValue "U_BPLId", 0, "3"
                Call FormItemEnabled
                If Left(Trim(Sbo_Company.UserName), 1) = "6" Then
            '        oForm01.Items("CardCode").Specific.Value = Trim(Sbo_Company.UserName)
                    oForm01.Items("CardCode").Specific.VALUE = Trim(Sbo_Company.UserName)
'                    oDS_PS_MM138H.setValue "U_CardCode", 0, Trim(Sbo_Company.UserName)
                    sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                    oRecordSet01.DoQuery sQry
                    oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'                    oDS_PS_MM138H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).Value)
                End If
                oForm01.Items("DocEntry").Click ct_Regular
                
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                Call Initialization
                Call FormClear
                Call FormItemEnabled
                'Add_MatrixRow 0, True
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.Freeze True
                If oCheck = "True" Then
                    Sbo_Application.ActivateMenuItem ("1281")
                    oForm01.Items("DocEntry").Specific.VALUE = oDocEntryNext
                    oForm01.Items("1").Click ct_Regular
                    
                    oCheck = "False"
                    oDocEntryNext = 0
                End If
                
                Call FormItemEnabled
                oForm01.Freeze False
            Case "1287": '// 복제
                oForm01.Freeze True
                Call FormClear
                oDS_PS_MM138H.setValue "Status", 0, "O"
                oDS_PS_MM138H.setValue "Canceled", 0, "N"
                
                Call FormItemEnabled

'                For i = 0 To oMat01.VisualRowCount - 1
'                    oMat01.FlushToDataSource
'                    oDS_PS_MM138L.setValue "DocEntry", i, ""
'                    oMat01.LoadFromDataSource
'                Next i
                oForm01.Freeze False
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
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
    '//디비데이터 소스 개체 할당
    Set oDS_PS_MM138H = oForm01.DataSources.DBDataSources("@PS_MM138H")
    Set oDS_PS_MM138L = oForm01.DataSources.DBDataSources("@PS_MM138L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
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
    
'    oMat01.Columns("OutGbn").ValidValues.Add "10", "제품SET기준"
'    oMat01.Columns("OutGbn").ValidValues.Add "20", "원재료기준"
        
'    oCombo.Select "N", psk_ByValue
    
    Set oCombo = oForm01.Items("IssueYN").Specific
    oCombo.ValidValues.Add "Y", "원재료불출완료"
    oCombo.ValidValues.Add "N", "원재료불출대상"
    oCombo.Select "N", psk_ByValue
    
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '//아이디별 사업장 세팅
'    Set oCombo = oForm01.Items("BPLId").Specific
'    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
'    oCombo.Select "1", psk_ByValue

    If Trim(oBPLId) = "" Then
        oBPLId = "3"
    End If
    
    If Trim(oCardCode) = "" Then
        oDocdate = Format(Now, "yyyymmdd")
    End If
    
    
    oDS_PS_MM138H.setValue "U_BPLId", 0, oBPLId '"3"
        
    oDS_PS_MM138H.setValue "U_DocDate", 0, oDocdate
    
    oDS_PS_MM138H.setValue "U_CardCode", 0, oCardCode
    oDS_PS_MM138H.setValue "U_CardName", 0, oCardName
    
    

    
    
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
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
    Dim DocEntry As String
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Items("DocDate").Click ct_Regular
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DocEntry").Enabled = False
        
        oForm01.Items("IssueYN").Enabled = True
        oForm01.Items("ItemCode").Enabled = True

        oMat01.Columns("OutQty").Editable = True
        oMat01.Columns("OutWt").Editable = True
        
        oForm01.Items("IssueYN").Specific.Select "N", psk_ByValue
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("DocEntry").Enabled = True

        oForm01.Items("ItemCode").Enabled = True
        
        oMat01.Columns("OutQty").Editable = False
        oMat01.Columns("OutWt").Editable = False
      
'        oMat01.Columns("ReStatus").Editable = False

        oForm01.Items("DocEntry").Click ct_Regular
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
    
        
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        sQry = "Select Count(*) From OCRD Where CardCode = '" & Trim(oDS_PS_MM138H.GetValue("U_CardCode", 0)) & "'"
        sQry = "select Count(*) From [@PS_MM095H] a Inner Join [@PS_MM095L] b On a.DocEntry = b.DocEntry"
        sQry = sQry + " Where Isnull(a.U_OIGEDoc,'') <> '' and Isnull(a.U_OIGNDoc,'') = '' "
        sQry = sQry + " And b.U_PP040Doc = '" & DocEntry & "'"
  
        oRecordSet01.DoQuery sQry
        '원재료 불출이 되었으면 선택못하게끔..
        If oRecordSet01.Fields(0).VALUE > 0 Then
            oForm01.Items("IssueYN").Enabled = False
        Else
            oForm01.Items("IssueYN").Enabled = True
        End If
        
        
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("CardCode").Enabled = False
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("Qty").Enabled = False
        
        oForm01.Items("ItemCode").Enabled = False
        
        oForm01.Items("DocEntry").Enabled = False

        oMat01.Columns("OutQty").Editable = True
        oMat01.Columns("OutWt").Editable = True
        
'                oMat01.Columns("ReStatus").Editable = True
        
        oForm01.EnableMenu ("1284"), True         '// 취소
    
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM138'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM138L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM138L.Offset = oRow
    oDS_PS_MM138L.setValue "U_LineNum", oRow, oRow + 1
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
    Dim sRow As Long
    Dim sSeq$
    Dim MDUseQty As Long, MDUseWt As Currency, MDUseUnWt As Currency, MUseWt As Currency
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    Select Case oUID
        Case "CardCode"
            sQry = "Select CardName From OCRD Where CardCode = '" & Trim(oDS_PS_MM138H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM138H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "ItmGrpCd"
            oForm01.Freeze True
            
            sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code ='M007' and U_Minor = '" & Trim(oForm01.Items("ItmGrpCd").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM138H.setValue "U_ItmGrpNm", 0, Trim(oRecordSet01.Fields(0).VALUE)
            
            oForm01.Freeze False
        Case "CpCode"
            oForm01.Freeze True


            sQry = "Select U_CpName From [@PS_PP001L] Where U_CpCode = '" & Trim(oDS_PS_MM138H.GetValue("U_CpCode", 0)) & "'"
            oRecordSet01.DoQuery sQry

            oDS_PS_MM138H.setValue "U_CpName", 0, Trim(oRecordSet01.Fields(0).VALUE)

            oForm01.Freeze False

            
        Case "ItemCode"
            oForm01.Freeze True
            
            oMat01.Clear
            
            oForm01.Items("Qty").Specific.VALUE = 0
            
            sQry = "Select ItemName, FrgnName, U_Size From OITM Where ItemCode = '" & Trim(oDS_PS_MM138H.GetValue("U_ItemCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM138H.setValue "U_ItemName", 0, Trim(oRecordSet01.Fields(0).VALUE)
            
            
            sQry = "EXEC [PS_MM138_01] '" & Trim(oDS_PS_MM138H.GetValue("U_BPLId", 0)) & "', '" & Trim(oDS_PS_MM138H.GetValue("U_ItemCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
                
            oMat01.FlushToDataSource
    
            i = 0
            Do Until oRecordSet01.EOF
                oDS_PS_MM138L.InsertRecord i
                oDS_PS_MM138L.setValue "U_LineNum", i, i + 1
                oDS_PS_MM138L.setValue "U_ItemCode", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)
                oDS_PS_MM138L.setValue "U_ItemName", i, Trim(oRecordSet01.Fields("ItemName").VALUE)
                oDS_PS_MM138L.setValue "U_Size", i, Trim(oRecordSet01.Fields("Size").VALUE)
                oDS_PS_MM138L.setValue "U_OutItmCd", i, Trim(oRecordSet01.Fields("MItemCod").VALUE)
                oDS_PS_MM138L.setValue "U_OutItmNm", i, Trim(oRecordSet01.Fields("MItemNam").VALUE)
'                oDS_PS_MM138L.setValue "U_OutGbn", i, Trim(oRecordSet01.Fields("OutGbn").VALUE)
                oDS_PS_MM138L.setValue "U_StdQty", i, Trim(oRecordSet01.Fields("StdQty").VALUE)
                oDS_PS_MM138L.setValue "U_BOMQty", i, Trim(oRecordSet01.Fields("BOMQty").VALUE)
                oDS_PS_MM138L.setValue "U_BOMWt", i, Trim(oRecordSet01.Fields("BOMWt").VALUE)
                oDS_PS_MM138L.setValue "U_InWhCd", i, Trim(oRecordSet01.Fields("InWhCd").VALUE)
                oDS_PS_MM138L.setValue "U_InWhNm", i, Trim(oRecordSet01.Fields("InWhNm").VALUE)
                i = i + 1
                oRecordSet01.MoveNext
            Loop
    
            oMat01.LoadFromDataSource
            
            oForm01.Freeze False
            'oMat01.FlushToDataSource
            
'        Case "Mat01"
'            If oCol = "OtDocLin" Then
'                oForm01.Freeze True
'                oMat01.FlushToDataSource
'
'                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("OtDocLin").Cells(oRow).Specific.VALUE) <> "" Then
'                    oMat01.FlushToDataSource
'                    Call Add_MatrixRow(oMat01.RowCount, False)
'                    oMat01.Columns("OtDocLin").Cells(oRow).Click ct_Regular
'                End If
'
'                sQry = "EXEC [PS_MM138_01] '" & Trim(oDS_PS_MM138H.GetValue("U_BPLId", 0)) & "', '" & Trim(oDS_PS_MM138H.GetValue("U_CardCode", 0)) & "', "
'                sQry = sQry & "'" & Trim(oDS_PS_MM138H.GetValue("U_OutDoc", 0)) & "'"
'                oRecordSet01.DoQuery sQry
'
'
'                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
'                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemName").VALUE)
'                oMat01.Columns("Size").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Size").VALUE)
'                oMat01.Columns("Mark").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Mark").VALUE)
'                oMat01.Columns("OutItmCd").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutItmCd").VALUE)
'                oMat01.Columns("OutItmNm").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutItmNm").VALUE)
'                oMat01.Columns("OutGbn").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_OutGbn").VALUE)
'                oMat01.Columns("InQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_InQty").VALUE)
'                oMat01.Columns("InWt").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_InWt").VALUE)
'                oMat01.Columns("OutDoc").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutDoc").VALUE)
'                oMat01.Columns("OutLine").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutLine").VALUE)
''                oMat01.Columns("ReStatus").Cells(oRow).Specific.Select "N"
'
'
'                oForm01.Freeze False
'
'            End If
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
        Case oDS_PS_MM138H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM138H.GetValue("U_CardCode", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM138H.GetValue("U_DocDate", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM138H.GetValue("U_CpCode", 0) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "외주거래처는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "전기일자는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "공정코드는 필수입력사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim MUseWt As Currency, MDUseWt As Currency, MUseQty As Currency, MDUseQty As Currency
    
    Dim OutDoc As String
    Dim OutLine As String
    Dim ItemCode As String
    
    Dim MM130Wt As Currency
    Dim MM152Wt As Currency
    Dim MM132Wt As Currency
    Dim OutWt   As Currency
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 Then
        If oDS_PS_MM138L.GetValue("U_ItemCode", 0) = "" Then
            ErrNum = 2
            GoTo MatrixSpaceLineDel_Error
        End If
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case Trim(oDS_PS_MM138L.GetValue("U_ItemCode", i)) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM138L.GetValue("U_Qty", i) = "" Or oDS_PS_MM138L.GetValue("U_Qty", i) = 0
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM138L.GetValue("U_OutWt", i) = 0 'And oDS_PS_MM138L.GetValue("U_ScrapWt", i) = 0
'                ErrNum = 5
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM138L.GetValue("U_Price", i) = 0
'                ErrNum = 6
'                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM138L.GetValue("U_LinTotal", i) = 0
'                ErrNum = 7
'                GoTo MatrixSpaceLineDel_Error
        End Select
        
        
    Next
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "첫라인에 반출문서-행 번호가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "반출문서-행 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 원재료사용중량은 원재료잔량중량보다 클수 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 출고중량은 0보다 커야만 합니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "단가는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "반출중량보다 입고중량이 큽니다. 확인하세요.", "E"
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
        If Trim(oDS_PS_MM138L.GetValue("U_ItemCode", i)) = "" Then
            oDS_PS_MM138L.RemoveRecord i
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'Public Function Add_PS_PP040(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
'On Error GoTo Add_PS_PP040_Error
'    Dim i&, j&
'    Dim ErrNum&, RowNum&
'    Dim sQry As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Dim lTypeCount As Integer
'
'    Dim BPLId$, ItemCode$, OutQty As Long, OutWt As Double
'    Dim PP040H_DocEntry$, PP030M_LineId$, PP030M_U_Sequence$, PP030M_U_CpCode$, PP030M_U_CpName$, DocType$, AutoKey&, DocDate$, OutGbn$, CpCode$
'    Dim iTemp01&, sTemp01$, ScrapWt As Currency
'
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
'    Sbo_Company.StartTransaction
'    oMat01.FlushToDataSource
'
'    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
'
'
'    If Trim(oDS_PS_MM138H.GetValue("U_OKYNC", 0)) = "C" Then
'        For i = 0 To oMat01.RowCount - 2
'            If Trim(oDS_PS_MM138L.GetValue("U_OtDocLin", i)) <> "" And Trim(oDS_PS_MM138L.GetValue("U_PP040Doc", i)) <> "" Then
'                OutQty = Trim(oDS_PS_MM138L.GetValue("U_OutQty", i))
'                OutWt = Trim(oDS_PS_MM138L.GetValue("U_OutWt", i))
'
'                sQry = "Update [@PS_PP040H] Set Status = 'C', Canceled = 'Y' Where DocEntry = '" & Trim(oDS_PS_MM138L.GetValue("U_PP040Doc", i)) & "'"
'                oRecordSet01.DoQuery sQry
'
''                sQry = "Update [@PS_MM130L] "
''                sQry = sQry & "Set U_InQty = IsNull(U_InQty, 0) - " & OutQty & ", U_InWt = IsNull(U_InWt, 0) - " & OutWt & " "
''                sQry = sQry & "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry "
''                sQry = sQry & "Where b.U_OutDoc = '" & Trim(oDS_PS_MM138L.GetValue("U_OutDoc", i)) & "' "
''                sQry = sQry & "And a.U_LineNum = '" & Trim(oDS_PS_MM138L.GetValue("U_OutLine", i)) & "' "
''                oRecordSet01.DoQuery sQry
'
''                sQry = "Update [@PS_MM138L] Set U_ReStatus = 'N' From [@PS_MM138H] a Where [@PS_MM138L].DocEntry = a.DocEntry "
''                sQry = sQry & "And a.Canceled = 'N' And a.U_OKYNC <> 'C' And [@PS_MM138L].DocEntry <> '" & Trim(oDS_PS_MM138L.GetValue("DocEntry", i)) & "' "
''                sQry = sQry & "And [@PS_MM138L].U_OtDocLin = '" & Trim(oDS_PS_MM138L.GetValue("U_OtDocLin", i)) & "' "
''                sQry = sQry & "And [@PS_MM138L].U_ReStatus = 'Y' "
''                oRecordSet01.DoQuery sQry
'
'
'                Call oDS_PS_MM138L.setValue("U_PP040Doc", i, "")
'            ElseIf Trim(oDS_PS_MM138L.GetValue("U_OtDocLin", i)) <> "" And Trim(oDS_PS_MM138L.GetValue("U_PP040Doc", i)) = "" Then
'            End If
'        Next i
'    ElseIf Trim(oDS_PS_MM138H.GetValue("U_OKYNC", 0)) = "Y" Then
'        For i = 0 To oMat01.RowCount - 2
'            If Trim(oDS_PS_MM138L.GetValue("U_OtDocLin", i)) <> "" And Trim(oDS_PS_MM138L.GetValue("U_PP040Doc", i)) <> "" Then
'            ElseIf Trim(oDS_PS_MM138L.GetValue("U_OtDocLin", i)) <> "" And Trim(oDS_PS_MM138L.GetValue("U_PP040Doc", i)) = "" Then
''                DocDate = Format(Trim(oDS_PS_MM138H.GetValue("U_DocDate", 0)), "YYYYMMDD")
'                DocDate = Trim(oDS_PS_MM138H.GetValue("U_DocDate", 0))
'
'                ItemCode = Trim(oDS_PS_MM138L.GetValue("U_ItemCode", i))
'                OutQty = Trim(oDS_PS_MM138L.GetValue("U_OutQty", i))  '외주수량
'                OutWt = Trim(oDS_PS_MM138L.GetValue("U_OutWt", i)) '외주중량
'                OutGbn = Trim(oDS_PS_MM138L.GetValue("U_OutGbn", i))
'
'                sQry = "EXEC [PS_MM138_02] '" & ItemCode & "'"
'                oRecordSet01.DoQuery sQry
'
'                If oRecordSet01.RecordCount = 0 Then
'                    ErrNum = 1
'                    GoTo Add_PS_PP040_Error
'                End If
'
'                lTypeCount = 0
'                Do Until oRecordSet01.EOF
'                    ReDim Preserve PS_PP040(lTypeCount)
'
'                    PS_PP040(lTypeCount).PP030HNo = ""
'                    PS_PP040(lTypeCount).PP030MNo = ""
'                    PS_PP040(lTypeCount).OrdNum = ItemCode
'                    PS_PP040(lTypeCount).Sequence = Trim(oRecordSet01.Fields("U_Sequence").VALUE)
'                    PS_PP040(lTypeCount).OrdGbn = Trim(oRecordSet01.Fields("U_OrdGbn").VALUE)
'                    PS_PP040(lTypeCount).ItemCode = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
'                    PS_PP040(lTypeCount).ItemName = Trim(oRecordSet01.Fields("U_ItemName").VALUE)
'                    PS_PP040(lTypeCount).CpCode = Trim(oRecordSet01.Fields("U_CpCode").VALUE)
'                    PS_PP040(lTypeCount).CpName = Trim(oRecordSet01.Fields("U_CpName").VALUE)
'                    PS_PP040(lTypeCount).Chk = "N"
'
'                    lTypeCount = lTypeCount + 1
'                    oRecordSet01.MoveNext
'                Loop
'
'                '// DocEntry
'                sQry = "Select AutoKey From [ONNM] Where ObjectCode = 'PS_PP040'"
'                oRecordSet01.DoQuery sQry
'                PP040H_DocEntry = Trim(oRecordSet01.Fields("AutoKey").VALUE)
'                AutoKey = PP040H_DocEntry + 1
'
'                ReDim Preserve PS_PP040DocEntry(i)
'                PS_PP040DocEntry(i).PP040DocEntry = PP040H_DocEntry
'
'                '//문서타입 하드코딩
'                If Trim(PS_PP040(0).OrdGbn) = "104" Or Trim(PS_PP040(0).OrdGbn) = "107" Then
'                    DocType = "20"
'                Else
'                    DocType = "10"
'                End If
'
'                '// Insert PS_PP040H
'                sQry = "INSERT INTO [@PS_PP040H]"
'                sQry = sQry & " ("
'                sQry = sQry & " DocEntry,"
'                sQry = sQry & " DocNum,"
'                sQry = sQry & " Period,"
'                sQry = sQry & " Series,"
'                sQry = sQry & " Object,"
'                sQry = sQry & " UserSign,"
'                sQry = sQry & " CreateDate,"
'                sQry = sQry & " CreateTime,"
'                sQry = sQry & " DataSource,"
'                sQry = sQry & " U_OrdType,"
'                sQry = sQry & " U_OrdGbn,"
'                sQry = sQry & " U_BPLId,"
'                sQry = sQry & " U_ItemCode,"
'                sQry = sQry & " U_ItemName,"
'                sQry = sQry & " U_OrdNum,"
'                sQry = sQry & " U_OrdSub1,"
'                sQry = sQry & " U_OrdSub2,"
'                sQry = sQry & " U_PP030HNo,"
'                sQry = sQry & " U_DocType,"
'                sQry = sQry & " Canceled,"
'                sQry = sQry & " U_DocDate"
'                sQry = sQry & " ) "
'
'                sQry = sQry & "VALUES("
'                sQry = sQry & "'" & PP040H_DocEntry & "',"
'                sQry = sQry & "'" & PP040H_DocEntry & "',"
'                sQry = sQry & "'22',"
'                sQry = sQry & "'-1',"
'                sQry = sQry & "'PS_PP040',"
'                sQry = sQry & "'1',"
'                sQry = sQry & "'" & Format(Date, "YYYYMMDD") & "',"
'                sQry = sQry & "'1000',"
'                sQry = sQry & "'I',"
'                sQry = sQry & "'30',"
'                sQry = sQry & "'" & PS_PP040(0).OrdGbn & "',"
'                sQry = sQry & "'" & BPLId & "',"
'                sQry = sQry & "'" & PS_PP040(0).ItemCode & "',"
'                sQry = sQry & "'" & PS_PP040(0).ItemName & "',"
'                sQry = sQry & "'" & ItemCode & "',"
'                sQry = sQry & "'00',"
'                sQry = sQry & "'000',"
'                sQry = sQry & "'" & PS_PP040(0).PP030HNo & "',"
'                sQry = sQry & "'" & DocType & "',"
'                sQry = sQry & "'N',"
'                sQry = sQry & "'" & DocDate & "'"
'                sQry = sQry & ")"
'                oRecordSet01.DoQuery sQry
'
'                '// Insert PS_PP040M
'                sQry = "INSERT INTO [@PS_PP040M]"
'                sQry = sQry & " ("
'                sQry = sQry & " DocEntry,"
'                sQry = sQry & " LineId,"
'                sQry = sQry & " VisOrder,"
'                sQry = sQry & " Object,"
'                sQry = sQry & " U_LineNum,"
'                sQry = sQry & " U_LineId,"
'                sQry = sQry & " U_WorkCode,"
'                sQry = sQry & " U_WorkName"
'                sQry = sQry & " ) "
'
'                sQry = sQry & "VALUES("
'                sQry = sQry & "'" & PP040H_DocEntry & "',"
'                sQry = sQry & "'1',"
'                sQry = sQry & "'0',"
'                sQry = sQry & "'PS_PP040',"
'                sQry = sQry & "'1',"
'                sQry = sQry & "'1',"
'                sQry = sQry & "'9999999',"
'                sQry = sQry & "'조정'"
'                sQry = sQry & ")"
'                oRecordSet01.DoQuery sQry
'
'                For j = 0 To (UBound(PS_PP040))
'                    If PS_PP040(j).Chk = "N" Then
'                        iTemp01 = j + 1
'                        sTemp01 = PS_PP040(j).PP030HNo & "-" & PS_PP040(j).PP030MNo
'
'                        '// Insert PS_PP040L
'                        sQry = "INSERT INTO [@PS_PP040L]"
'                        sQry = sQry & " ("
'                        sQry = sQry & " DocEntry,"
'                        sQry = sQry & " LineId,"
'                        sQry = sQry & " VisOrder,"
'                        sQry = sQry & " Object,"
'                        sQry = sQry & " U_LineNum,"
'                        sQry = sQry & " U_LineId,"
'                        sQry = sQry & " U_OrdMgNum,"
'                        sQry = sQry & " U_Sequence,"
'                        sQry = sQry & " U_CpCode,"
'                        sQry = sQry & " U_CpName,"
'                        sQry = sQry & " U_OrdGbn,"
'                        sQry = sQry & " U_BPLId,"
'                        sQry = sQry & " U_ItemCode,"
'                        sQry = sQry & " U_ItemName,"
'                        sQry = sQry & " U_OrdNum,"
'                        sQry = sQry & " U_OrdSub1,"
'                        sQry = sQry & " U_OrdSub2,"
'                        sQry = sQry & " U_PP030HNo,"
'                        sQry = sQry & " U_PP030MNo,"
'                        sQry = sQry & " U_PQty,"
'                        sQry = sQry & " U_PWeight,"
'                        sQry = sQry & " U_YQty,"
'                        sQry = sQry & " U_YWeight,"
'                        sQry = sQry & " U_NQty,"
'                        sQry = sQry & " U_NWeight,"
'                        sQry = sQry & " U_PSum"
'                        sQry = sQry & " ) "
'
'                        sQry = sQry & "VALUES("
'                        sQry = sQry & "'" & PP040H_DocEntry & "',"
'                        sQry = sQry & "'" & iTemp01 & "',"
'                        sQry = sQry & "'" & j & "',"
'                        sQry = sQry & "'PS_PP040',"
'                        sQry = sQry & "'" & iTemp01 & "',"
'                        sQry = sQry & "'" & iTemp01 & "',"
'                        sQry = sQry & "'" & sTemp01 & "',"
'                        sQry = sQry & "'" & PS_PP040(j).Sequence & "',"
'                        sQry = sQry & "'" & PS_PP040(j).CpCode & "',"
'                        sQry = sQry & "'" & PS_PP040(j).CpName & "',"
'                        sQry = sQry & "'" & PS_PP040(j).OrdGbn & "',"
'                        sQry = sQry & "'" & BPLId & "',"
'                        sQry = sQry & "'" & PS_PP040(j).ItemCode & "',"
'                        sQry = sQry & "'" & PS_PP040(j).ItemName & "',"
'                        sQry = sQry & "'" & ItemCode & "',"
'                        sQry = sQry & "'000',"
'                        sQry = sQry & "'00',"
'                        sQry = sQry & "'" & PS_PP040(j).PP030HNo & "',"
'                        sQry = sQry & "'" & PS_PP040(j).PP030MNo & "',"
'                        sQry = sQry & "'" & OutQty & "',"
'                        sQry = sQry & "'" & OutWt & "',"
'                        sQry = sQry & "'" & OutQty & "',"
'                        sQry = sQry & "'" & OutWt & "',"
'                        sQry = sQry & "'" & 0 & "',"
'                        sQry = sQry & "'" & 0 & "',"
'                        sQry = sQry & "'" & 0 & "'"
'                        sQry = sQry & ")"
'                        oRecordSet01.DoQuery sQry
'
'                        '// Insert PS_PP040N
'                        sQry = "INSERT INTO [@PS_PP040N]"
'                        sQry = sQry & " ("
'                        sQry = sQry & " DocEntry,"
'                        sQry = sQry & " LineId,"
'                        sQry = sQry & " VisOrder,"
'                        sQry = sQry & " Object,"
'                        sQry = sQry & " U_LineNum,"
'                        sQry = sQry & " U_LineId,"
'                        sQry = sQry & " U_OrdMgNum,"
'                        sQry = sQry & " U_CpCode,"
'                        sQry = sQry & " U_CpName"
'                        sQry = sQry & " ) "
'
'                        sQry = sQry & "VALUES("
'                        sQry = sQry & "'" & PP040H_DocEntry & "',"
'                        sQry = sQry & "'" & iTemp01 & "',"
'                        sQry = sQry & "'" & j & "',"
'                        sQry = sQry & "'PS_PP040',"
'                        sQry = sQry & "'" & iTemp01 & "',"
'                        sQry = sQry & "'" & iTemp01 & "',"
'                        sQry = sQry & "'" & sTemp01 & "',"
'                        sQry = sQry & "'" & PS_PP040(j).CpCode & "',"
'                        sQry = sQry & "'" & PS_PP040(j).CpName & "'"
'                        sQry = sQry & ")"
'                        oRecordSet01.DoQuery sQry
'
'                        PS_PP040(j).Chk = "Y"
'                    End If
'                Next j
'                sQry = "Update [ONNM] Set AutoKey = '" & AutoKey & "' Where ObjectCode = 'PS_PP040'"
'                oRecordSet01.DoQuery sQry
'
''                sQry = "Update [@PS_MM130L] "
''                sQry = sQry & "Set U_InQty = IsNull(U_InQty, 0) + " & OutQty & ", U_InWt = IsNull(U_InWt, 0) + " & OutWt & " "
''                sQry = sQry & "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry "
''                sQry = sQry & "Where b.U_OutDoc = '" & Trim(oDS_PS_MM138L.GetValue("U_OutDoc", i)) & "' "
''                sQry = sQry & "And a.U_LineNum = '" & Trim(oDS_PS_MM138L.GetValue("U_OutLine", i)) & "' "
''                oRecordSet01.DoQuery sQry
'
'                Call oDS_PS_MM138L.setValue("U_PP040Doc", i, PS_PP040DocEntry(i).PP040DocEntry)
'            End If
'        Next i
'    End If
'
'    oMat01.LoadFromDataSource
'    MDC_Com.MDC_GF_Message "외주업체 출고등록 완료!", "S"
'
'    Set oRecordSet01 = Nothing
'    Add_PS_PP040 = True
'    Exit Function
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Add_PS_PP040_Error:
'    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
'    Add_PS_PP040 = False
'    Set oRecordSet01 = Nothing
'    If ErrNum = 1 Then
'        MDC_Com.MDC_GF_Message "작업지시가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
'    Else
'        MDC_Com.MDC_GF_Message "Add_PS_PP040_Error:" & Err.Number & " - " & Err.Description, "E"
'    End If
'End Function

Private Sub Print_Report01()
On Error GoTo Print_Report01_Error
    Dim ErrNum&
    Dim DocNum                  As String
    Dim WinTitle                As String
    Dim ReportName              As String
    Dim sQry                    As String
    Dim oRecordSet01            As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call ConnectODBC
    
   
    DocNum = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
    If DocNum = "" Then
        ErrNum = 1
        GoTo Print_Report01_Error
    End If
        
    WinTitle = "거래명세표 [PS_MM138_01]"
    ReportName = "PS_MM138_01.rpt"
       
    '//Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '//SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '//조회조건문
    sQry = "EXEC [PS_MM138_03] '" & DocNum & "'"
    oRecordSet01.DoQuery sQry
    If oRecordSet01.RecordCount = 0 Then
        ErrNum = 2
        GoTo Print_Report01_Error
    End If
    
    '//CR Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report01_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Or ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "승인되지 않은 문서는 거래명세표를 출력할 수 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Function PS_MM138_DI_API01() As Boolean
On Error GoTo PS_MM138_DI_API01_Error
    PS_MM138_DI_API01 = True
    Dim i, j As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim ResultDocNum As Long
    Dim BaseLot As String
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
        
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenEntry)
    oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    
    oDIObject.Lines.ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    oDIObject.Lines.WarehouseCode = "10" + Trim(oForm01.Items("BPLId").Specific.VALUE)
    oDIObject.Lines.Quantity = Trim(oForm01.Items("Qty").Specific.VALUE)
            
    oDIObject.UserFields("U_OrdTyp").VALUE = "30"
    oDIObject.UserFields("U_CardCode").VALUE = Trim(oForm01.Items("CardCode").Specific.VALUE) '//입고취소 문서번호
    oDIObject.UserFields("U_CardName").VALUE = Trim(oForm01.Items("CardName").Specific.VALUE) '//입고취소 문서번호
    
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        oForm01.Items("OIGNNo").Specific.VALUE = ResultDocNum
        Call oDS_PS_MM138H.setValue("U_OIGNNo", 0, ResultDocNum)
    Else
        GoTo PS_MM138_DI_API01_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    
    
    Set oDIObject = Nothing
    Exit Function
PS_MM138_DI_API01_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_MM138_DI_API01 = False
    Set oDIObject = Nothing
    Exit Function
PS_MM138_DI_API01_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_MM138_DI_API01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_MM138_DI_API01 = False
End Function


Private Function PS_MM138_DI_API02() As Boolean
On Error GoTo PS_MM138_DI_API02_Error
    PS_MM138_DI_API02 = True
    
    Dim i, j, K As Long
    Dim oDIObject As SAPbobsCOM.Documents
    Dim RetVal As Long
    Dim ResultDocNum As Long
    Dim BaseLot As String
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    
    Set oDIObject = Sbo_Company.GetBusinessObject(oInventoryGenExit)
    oDIObject.DocDate = Format(oForm01.Items("DocDate").Specific.VALUE, "&&&&-&&-&&")
    oDIObject.UserFields("U_CancDoc").VALUE = Trim(oForm01.Items("OIGNNo").Specific.VALUE)
    oDIObject.UserFields("U_OrdTyp").VALUE = "30"
    oDIObject.UserFields("U_CardCode").VALUE = Trim(oForm01.Items("CardCode").Specific.VALUE) '//거래처코드
    oDIObject.UserFields("U_CardName").VALUE = Trim(oForm01.Items("CardName").Specific.VALUE) '//거래처명
    
    oDIObject.Lines.ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE) '//품목코드
    oDIObject.Lines.WarehouseCode = "10" + Trim(oForm01.Items("BPLId").Specific.VALUE)
    oDIObject.Lines.Quantity = Trim(oForm01.Items("Qty").Specific.VALUE) '//수량
    
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        oForm01.Items("OIGENo").Specific.VALUE = ResultDocNum
        Call oDS_PS_MM138H.setValue("U_OIGENo", 0, ResultDocNum)
    Else
        GoTo PS_MM138_DI_API02_DI_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    
    
    Set oDIObject = Nothing
    Set oDIObject = Nothing
    Exit Function
PS_MM138_DI_API02_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_MM138_DI_API02 = False
    Set oDIObject = Nothing
    Exit Function
PS_MM138_DI_API02_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage "PS_MM138_DI_API02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_MM138_DI_API02 = False
End Function

Private Sub SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo SetDocument_Error
    If (oFromDocEntry01 = "") Then

    Else
        oForm01.Mode = fm_FIND_MODE
        Call FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
