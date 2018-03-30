VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP005.cls
'//  Module         : PP
'//  Description    : 제품 원재료관계등록
'//  FormType       : PS_PP005H
'//  Create Date    : 2010.10.20
'//  Creator        : Lee Byong Gak
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_PP005H As SAPbouiCOM.DBDataSource          '등록헤더

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP005.srf")
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
    
    oFormUniqueID01 = "PS_PP005_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
       
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
 '   oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
'    Call Initialization
'    Call PS_PP005_CF_ChooseFromList
'    Call PS_PP005_FormItemEnabled
'    Call FormClear
    Call Add_MatrixRow(0, True)
                  
    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1286"), True         '// 닫기
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    oForm01.EnableMenu ("1282"), True
    
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

'// 데이타 insert
Public Function Add_PurchaseDemand(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo Add_PurchaseDemand_Error
    Dim i As Integer
    Dim sQry As String
    Dim sQry1 As String
    Dim ReturnValue As Integer
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, DocNum$, BPLID$, LineNum$, WorkTime As Currency
    Dim UnWeight As Double
    Dim ItemCod1$, ItemNam1$, ItemCod2$, ItemNam2$, Indate$
                
    ItemCod1 = Trim(oForm01.Items("ItemCod1").Specific.VALUE)
    ItemNam1 = Trim(oForm01.Items("ItemNam1").Specific.VALUE)
    ItemCod2 = Trim(oForm01.Items("ItemCod2").Specific.VALUE)
    ItemNam2 = Trim(oForm01.Items("ItemNam2").Specific.VALUE)
    UnWeight = Trim(oForm01.Items("UnWeight").Specific.VALUE)
    Indate = Format(Date, "yymmdd")                                                              '오늘날짜를 넣어주기 위한 선언부분
      
    sQry1 = "Select U_ItemCod1, U_ItemCod2 From [@PS_PP005H] Where U_ItemCod1 ='" & ItemCod1 & "' AND U_ItemCod2 = '" & ItemCod2 & "'"
    
    RecordSet01.DoQuery sQry1
    If (RecordSet01.RecordCount > 0) Then
        MDC_Com.MDC_GF_Message "기존자료가 존재합니다.:" & Err.Number & " - " & Err.Description, "E"
'        ReturnValue = Sbo_Application.MessageBox("기존자료가 존재합니다.", 1, "&확인", "&취소")   '오류시 팝업창에 Error표시
        Set RecordSet01 = Nothing
        Exit Function
    End If
        
    If UnWeight <= 0 Then
        MDC_Com.MDC_GF_Message "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.:" & Err.Number & " - " & Err.Description, "E"
        Set RecordSet01 = Nothing
        Exit Function
    End If
    
    sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_PP005H]"
    RecordSet01.DoQuery sQry
'        If Trim(RecordSet01.Fields(0).Value) = 0 Then
'            DocEntry = Left(DocDate, 6) + "0001"
'        Else
            DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
'        End If

'        ItemCod1 = Trim(oDS_PS_PP005H.GetValue("U_ItemCod1", i))
'        CntcName = Trim(oDS_PS_PP060H.GetValue("U_CntcName", i))
'        ItmBSort = Trim(oDS_PS_PP060H.GetValue("U_ItmBsort", i))
'        WorkNote = Trim(oDS_PS_PP060H.GetValue("U_WorkNote", i))
'        WorkTime = Trim(oDS_PS_PP060H.GetValue("U_WorkTime", i))

        sQry = "INSERT INTO [@PS_PP005H]"
        sQry = sQry & " ("
        sQry = sQry & " DocEntry,"
        sQry = sQry & " DocNum,"
        sQry = sQry & " U_ItemCod1,"
        sQry = sQry & " U_ItemNam1,"
        sQry = sQry & " U_ItemCod2,"
        sQry = sQry & " U_ItemNam2,"
        sQry = sQry & " U_UnWeight,"
        sQry = sQry & " U_InDate"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("
        sQry = sQry & DocEntry & ","
        sQry = sQry & DocEntry & ","
        sQry = sQry & "'" & ItemCod1 & "',"
        sQry = sQry & "'" & ItemNam1 & "',"
        sQry = sQry & "'" & ItemCod2 & "',"
        sQry = sQry & "'" & ItemNam2 & "',"
        sQry = sQry & UnWeight & ","
        sQry = sQry & "'" & Indate & "'"
        sQry = sQry & ")"
        RecordSet01.DoQuery sQry
'        Next
    
    MDC_Com.MDC_GF_Message "제품코드 및 원자재코드 정상등록!", "S"
    
    Set RecordSet01 = Nothing
    Add_PurchaseDemand = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_PurchaseDemand_Error:
    Add_PurchaseDemand = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Add_PurchaseDemand_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

'// 데이타 Update
Public Function UpdateData() As Boolean
On Error GoTo UpdateData_Error
    Dim i As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, ItemCode$, ItemName$, Qty&, Weight As Currency, Price As Currency, LinTotal As Currency
    Dim ItemCod2$, ItemNam2
    Dim Chk$, MoDate$
    Dim UnWeight As Double
    
    MoDate = Format(Date, "yymmdd")                       '오늘날짜를 불러오기 위한 변수선언
    
    
    oMat01.FlushToDataSource
        
    For i = 1 To oMat01.RowCount
        Chk = oMat01.Columns("Chk").Cells(i).Specific.Checked
         
        If Chk = True Then
            UnWeight = oMat01.Columns("UnWeight").Cells(i).Specific.VALUE
            
            If UnWeight <= 0 Then
                MDC_Com.MDC_GF_Message "단조품은 1, 중량단중은 개당 단중을 입력하여야 합니다.:" & Err.Number & " - " & Err.Description, "E"
                Exit Function
            End If
        End If
    Next
    
    For i = 1 To oMat01.RowCount
        Chk = oMat01.Columns("Chk").Cells(i).Specific.Checked
         
        If Chk = True Then
           DocEntry = oMat01.Columns("DocEntry").Cells(i).Specific.VALUE
           ItemCod2 = oMat01.Columns("ItemCod2").Cells(i).Specific.VALUE
           ItemNam2 = oMat01.Columns("ItemNam2").Cells(i).Specific.VALUE
           UnWeight = oMat01.Columns("UnWeight").Cells(i).Specific.VALUE
            
           sQry = "Update [@PS_PP005H] set "
           sQry = sQry & " U_ItemCod2   = '" & ItemCod2 & "',"
           sQry = sQry & " U_ItemNam2   = '" & ItemNam2 & "',"
           sQry = sQry & " U_UnWeight   = " & UnWeight & ","
           sQry = sQry & " U_MoDate     = '" & MoDate & "'"                  '오늘날짜 가져오기
            
           sQry = sQry & " Where DocEntry = '" & DocEntry & "'"
           RecordSet01.DoQuery sQry
        End If
    Next

    MDC_Com.MDC_GF_Message "원자재코드 수정완료!", "S"

    Set RecordSet01 = Nothing
    UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
UpdateData_Error:
    UpdateData = False
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Sub DeleteData()
On Error GoTo DeleteData_Error
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry$, ItemCode$, ItemName$, Qty&, Weight As Currency, Price As Currency, LinTotal As Currency
    Dim ItemCod2$, ItemNam2
    Dim Chk$, MoDate$
    
    oMat01.FlushToDataSource
    
     For i = 1 To oMat01.RowCount
        Chk = oMat01.Columns("Chk").Cells(i).Specific.Checked
         
         If Chk = True Then
            DocEntry = oMat01.Columns("DocEntry").Cells(i).Specific.VALUE
            ItemCod2 = oMat01.Columns("ItemCod2").Cells(i).Specific.VALUE
            ItemNam2 = oMat01.Columns("ItemNam2").Cells(i).Specific.VALUE
            
            sQry = "Delete From [@PS_PP005H] where DocEntry = '" & DocEntry & "'"
            oRecordSet01.DoQuery sQry
        End If
  
    Next i
     
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    Call Add_MatrixRow(0, True)
    Call LoadData
    Exit Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
DeleteData_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "DeleteData_Error:" & Err.Number & " - " & Err.Description, "E"

End Sub

'//조회데이타 가져오기
Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ItemCod1$, ItemCod2$, ItmBsort$, ItmMsort$

    ItemCod1 = Trim(oForm01.Items("ItemCod1").Specific.VALUE)
    ItemCod2 = Trim(oForm01.Items("ItemCod2").Specific.VALUE)
    ItmBsort = Trim(oForm01.Items("ItmBSort").Specific.VALUE)
    ItmMsort = Trim(oForm01.Items("ItmMSort").Specific.VALUE)

    If ItemCod1 = "" Then ItemCod1 = "%"
    If ItemCod2 = "" Then ItemCod2 = "%"
    If ItmMsort = "" Then ItmMsort = "%"
    If ItmBsort = "" Then ItmBsort = "%"

   sQry = "EXEC [PS_PP005_01] '" & ItmBsort & "','" & ItmMsort & "','" & ItemCod1 & "','" & ItemCod2 & "'"
   
    oRecordSet01.DoQuery sQry

    oMat01.Clear
    oDS_PS_PP005H.Clear
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If

    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)

    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP005H.Size Then
            oDS_PS_PP005H.InsertRecord (i)
        End If

        oMat01.AddRow
        oDS_PS_PP005H.Offset = i

        oDS_PS_PP005H.setValue "DocEntry", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)
        oDS_PS_PP005H.setValue "DocNum", i, i + 1
        oDS_PS_PP005H.setValue "U_ItemCod1", i, Trim(oRecordSet01.Fields("U_ItemCod1").VALUE)
        oDS_PS_PP005H.setValue "U_ItemNam1", i, Trim(oRecordSet01.Fields("U_ItemNam1").VALUE)
        oDS_PS_PP005H.setValue "U_ItmMSort", i, Trim(oRecordSet01.Fields("U_ItmMSort").VALUE)
'        oMat01.Columns("ItmMsort").Cells(i + 1).Specific.String = Trim(oRecordSet01.Fields("U_ItmMsort").Value)
        oDS_PS_PP005H.setValue "U_ItemCod2", i, Trim(oRecordSet01.Fields("U_ItemCod2").VALUE)
        oDS_PS_PP005H.setValue "U_ItemNam2", i, Trim(oRecordSet01.Fields("U_ItemNam2").VALUE)
        oDS_PS_PP005H.setValue "U_UnWeight", i, Trim(oRecordSet01.Fields("U_UnWeight").VALUE)
        oDS_PS_PP005H.setValue "U_InDate", i, Format(Trim(oRecordSet01.Fields("U_InDate").VALUE), "YYYYMMDD")
        oDS_PS_PP005H.setValue "U_MoDate", i, Format(Trim(oRecordSet01.Fields("U_MoDate").VALUE), "YYYYMMDD")
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    
    oForm01.Freeze False

    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
       
    Exit Sub
    
LoadData_Error:
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
    
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

On Error GoTo Raise_ItemEvent_Error
    
Dim i As Integer
Dim sQry  As String
Dim oRecordSet01 As SAPbobsCOM.Recordset

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1                               '버튼 클릭시 발생하는 Event
                If pval.ItemUID = "Btn01" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        If oForm01.Mode = fm_ADD_MODE Then
                        If pval.ItemUID = "Btn01" Then                '저장 버튼클릭
                                If oForm01.Mode = fm_ADD_MODE Then
                                    If HeaderSpaceLineDel = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                    
                                    If Add_PurchaseDemand(pval) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                    
                                    oMat01.Clear
                                    oMat01.FlushToDataSource
                                    oMat01.LoadFromDataSource
                                    Add_MatrixRow 0, True
                    
'                        Call Delete_EmptyRow
                                    oLast_Mode = oForm01.Mode
'                                ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'                                    If Updatedata(pval) = False Then
'                                        BubbleEvent = False
'                                        Exit Sub
'                                    End If
'                                    Call LoadData
                                End If
                            End If
                        End If
                        oLast_Mode = oForm01.Mode
                    End If
                ElseIf pval.ItemUID = "Btn02" Then                     '내역조회
                    Call LoadData
                ElseIf pval.ItemUID = "Btn03" Then
                    oForm01.Mode = fm_OK_MODE '/fm_VIEW_MODE
                    Call DeleteData
                ElseIf pval.ItemUID = "Btn04" Then
                    oForm01.Mode = fm_UPDATE_MODE
                    Call UpdateData
                End If
          Exit Sub
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2                         '탭키를 눌렀을 때 발생하는 Event

' 질의 관리자 창을 사용할 때... 선언부분은 FlushTo ItemValue에서 선언

                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "ItemCod1" Then
                        If oForm01.Items("ItemCod1").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")                '질의 관리자 사용시(탭키를 눌렀을 때)
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ItemCod2" Then
                        If oForm01.Items("ItemCod2").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ColUID = "ItemCod2" Then                             ' Matrix에 질의 관리자 사용시 선언
                        If oMat01.Columns("ItemCod2").Cells(pval.Row).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
                
'et_COMBO_SELECT ////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                
            
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE//////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    Add_MatrixRow oMat01.RowCount, False
                    oLast_Mode = 0
'                    If oForm01.Mode = fm_OK_MODE And oLast_Mode = fm_UPDATE_MODE Then
'
'                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5

                If pval.ItemUID = "ItmBSort" Then
                    Dim sCount&, sSeq&
                    sCount = Trim(oForm01.Items("ItmMSort").Specific.ValidValues.Count)
                    sSeq = sCount
                    For i = 1 To sCount
                        oForm01.Items("ItmMSort").Specific.ValidValues.Remove sSeq - 1, psk_Index
                        sSeq = sSeq - 1
                    Next i

                    '//중분류
                    sQry = "SELECT U_Code, U_CodeName From [@PSH_ITMMSORT] Where U_rCode = '" & Trim(oForm01.Items("ItmBSort").Specific.VALUE) & "' Order by Code"
                    oRecordSet01.DoQuery sQry
'                    oForm01.Items("ItmMSort").Specific.ValidValues.Add "", ""
                    Do Until oRecordSet01.EOF
                        oForm01.Items("ItmMSort").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
                        oRecordSet01.MoveNext
                    Loop
'                    If oRecordSet01.RecordCount <> 0 Then
                     oForm01.Items("ItmMSort").Specific.Select 0, psk_Index
'                    Else
'                        oForm01.Items("ItmMSort").Specific.Select "", psk_ByValue
'                    End If
                ElseIf pval.ItemUID = "Mat01" Then
'''                    If oForm01.Mode = fm_ADD_MODE Then
'''                    Else
'''                        oForm01.Mode = fm_UPDATE_MODE
'''                        Call LoadCaption
'''                    End If
                End If

            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10       ' 필드의 값이 바뀌었을 때 동작하는 부분
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "ItemCod1" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "ItemCod2" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ColUID = "ItemCod2" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                End If
                
'                If pval.ItemChanged = True Then
'                    If pval.ItemUID = "CardCode" Then
'                        FlushToItemValue pval.ItemUID
'                    ElseIf pval.ItemUID = "CntcCode" Then
'                        FlushToItemValue pval.ItemUID
'                    ElseIf pval.ItemUID = "Mat01" Then
'                        If pval.ColUID = "GADocLin" Then
'                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
'                        End If
'                    End If
'                End If
                
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
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
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
'            Case "1284": '취소
'                FormItemEnabled
'                oForm01.Items("DocNum").Click ct_Regular
            Case "1286": '닫기
            Case "1293": '행삭제
                
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_PP005H.RemoveRecord oDS_PS_PP005H.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
                    If oMat01.Columns("PQDocNum").Cells(oMat01.RowCount).Specific.VALUE <> "" Then
                        Call Add_MatrixRow(oMat01.RowCount, False)
                    End If
                End If
            
            Case "1281": '찾기
                FormItemEnabled
                'oForm01.Items("DocNum").Click ct_Regular
                
            Case "1282": '추가
'                oForm01.Items("ItemCod1").Specific.Value = ""
'                oForm01.Items("ItemCod2").Specific.Value = ""
                oMat01.Clear
                oDS_PS_PP005H.Clear
'                oDS_PS_PP005H.GetValue("U_ItemCod1",0)
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("PQDocNum").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_PP005H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
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

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
'                If Add_oPurchaseOrders(2) = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                Else
'                    Call Delete_EmptyRow
'                End If
                
            Case et_FORM_DATA_UPDATE: '//35
                If oLast_Mode = fm_UPDATE_MODE Then
'                    If Update_oPurchaseOrders(2) = False Then
'                        oLast_Mode = 0
'                        BubbleEvent = False
'                        Exit Sub
'                    Else
'                        oLast_Mode = 0
'                        Call Delete_EmptyRow
'                    End If
                ElseIf oLast_Mode = 101 Then '//취소
'                    If Cancel_oPurchaseOrders(2) = False Then
'                        oLast_Mode = 0
'                        BubbleEvent = False
'                        Exit Sub
'                    Else
'                        oLast_Mode = 0
'                    End If
                ElseIf oLast_Mode = 102 Then '//닫기
'                    If Close_oPurchaseOrders(2) = False Then
'                        oLast_Mode = 0
'                        BubbleEvent = False
'                        Exit Sub
'                    Else
'                        oLast_Mode = 0
'                    End If
                End If
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
'    oForm01.DataSources.DataTables.Add ("PS_PP005")
    
    Set oDS_PS_PP005H = oForm01.DataSources.DBDataSources("@PS_PP005H")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    

' 화면 호출시 등록일자에 오늘날짜 뿌려주기
'
    Call oForm01.DataSources.UserDataSources.Add("UnWeight", dt_PERCENT, 10)
    Call oForm01.Items("UnWeight").Specific.DataBind.SetBound(True, "", "UnWeight")
    
'   Call oForm01.DataSources.UserDataSources.Add("InDateFr", dt_DATE, 8)
'   Call oForm01.DataSources.UserDataSources.Add("InDateTo", dt_DATE, 8)
'   Call oForm01.Items("InDateFr").Specific.DataBind.SetBound(True, "", "InDateFr")
'   Call oForm01.Items("InDateTo").Specific.DataBind.SetBound(True, "", "InDateTo")
'
'   oForm01.Items("InDateFr").Specific.Value = Format(Now, "yyyymmdd")
'   oForm01.Items("InDateTo").Specific.Value = Format(Now, "yyyymmdd")


'    oDS_PS_PP005H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
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
   
    '// 대분류
    Set oCombo = oForm01.Items("ItmBSort").Specific
    sQry = "SELECT Code,Name From [@PSH_ITMBSORT] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    oCombo.Select 0, psk_Index

    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub CF_ChooseFromList()
On Error GoTo PS_PP005_CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP005_CF_ChooseFromList_Error:
    MDC_Com.MDC_GF_Message "CF_ChooseFromList_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
'    If (oForm01.Mode = fm_ADD_MODE) Then
'        oForm01.Items("DocNum").Enabled = False
'        oForm01.Items("DocDate").Enabled = True
'    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("DocNum").Enabled = True
'        oForm01.Items("DocDate").Enabled = True
'    ElseIf (oForm01.Mode = fm_OK_MODE) Then
'        oForm01.Items("DocNum").Enabled = False
'        oForm01.Items("DocDate").Enabled = False
'    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP005'", "")
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
'    If RowIserted = False Then '//행추가여부
'        oDS_PS_PP005H.InsertRecord (oRow)
'    End If
'    oMat01.AddRow
'    oDS_PS_PP005H.Offset = oRow
'    oDS_PS_PP005H.setValue "U_DocNum", oRow, oRow + 1
'    oDS_PS_PP005H.setValue "U_Chk", oRow, oRow + 1
'    oDS_PS_PP005H.setValue "U_ItemCod1", oRow, oRow + 1
'    oDS_PS_PP005H.setValue "U_ItemNam1", oRow, oRow + 1
'    oDS_PS_PP005H.setValue "U_ItemCod2", oRow, oRow + 1
''    oDS_PS_PP005H.setValue "U_ItmMsort", oRow, oRow + 1
'''    oDS_OITM.setValue "U_ItmMsort", oRow, oRow + 1
'    oDS_PS_PP005H.setValue "U_ItemNam2", oRow, oRow + 1
'    oDS_PS_PP005H.setValue "U_InDate", oRow, oRow + 1
'    oDS_PS_PP005H.setValue "U_MoDate", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

' 코드에 대한 품명을 뿌려주기 위한 선언부분

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sRow As Long
    Dim sSeq$
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sRow = oRow
     
    Select Case oUID
        Case "ItemCod1"
            sQry = "Select ItemName, ItmMSort = U_ItmMsort  From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCod1").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("ItemNam1").Specific.String = Trim(oRecordSet01.Fields("ItemName").VALUE)
            
            Set oCombo = oForm01.Items("ItmMSort").Specific
            oCombo.Select Trim(oRecordSet01.Fields("ItmMSort").VALUE), psk_ByValue
            'oDS_PS_PP005H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
            
        Case "ItemCod2"
            sQry = "Select ItemName  From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCod2").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.Items("ItemNam2").Specific.String = Trim(oRecordSet01.Fields("ItemName").VALUE)
            
    End Select
    
' Matrix 필드에 질의 응답 창 띄워주기
   Select Case oCol
        Case "ItemCod2"
            sQry = "Select FrgnName  From OITM Where ItemCode = '" & Trim(oMat01.Columns("ItemCod2").Cells(oRow).Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oMat01.Columns("ItemNam2").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("FrgnName").VALUE)
   End Select
   
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
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

' 저장버튼 클릭시 필수입력 필드에 값이 있는지를 Check 한다.

    Select Case True
        Case Trim(oForm01.Items("ItemCod1").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ItemNam1").Specific.VALUE) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ItemCod2").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ItemNam2").Specific.VALUE) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True

Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "제품코드는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "제품코드명은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "원자재코드는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "원자재명은 필수입력사항입니다. 확인하세요.", "E"
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

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

'Sub Delete_EmptyRow()
'On Error GoTo Delete_EmptyRow_Error
'    Dim i&
'
'    oMat01.FlushToDataSource
'
'    For i = 0 To oMat01.VisualRowCount - 1
'        If Trim(oDS_PS_PP005L.GetValue("U_ItemCode", i)) = "" Then
'            oDS_PS_PP005L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
'        End If
'    Next i
'
'    oMat01.LoadFromDataSource
'    Exit Sub
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Delete_EmptyRow_Error:
'    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
'End Sub
