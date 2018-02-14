VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM010.cls
'//  Module         : MM
'//  Description    : 구매견적
'//  FormType       : PS_MM010
'//  Create Date    : 2010.08.23
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************

'//U_Gubun 통합구매 구분값 1: 초기값, 2: 통합구매에서 품의생성, 3:홀딩스에서 품의생성

Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PS_MM010H  As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM010L  As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oSeq&
Private oCount&
Private oPurchase  As String
Private oPQType    As String

Dim hOpen          As Long
Dim hConnection    As Long
Dim dwType         As Long

Const ASCII        As Long = FTP_TRANSFER_TYPE_ASCII
Const BINARY       As Long = FTP_TRANSFER_TYPE_BINARY
Const FTP_SERVER   As String = "192.1.11.3"  '"192.1.11.7" '7번은 테스트용
Const FTP_USER     As String = "ftpadm"
Const FTP_PASSWORD As String = "psc1004"

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional sDocNum As String, Optional SBPLID As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM010.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_MM010_" & GetTotalFormsCount
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
        Call oForm01.Items("BPLId").Specific.Select(SBPLID, psk_ByValue)
                
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
    Dim Comments$
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    Dim sQry            As String
    Dim ItmBsort        As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FileName As String
    Dim Dir As String
    
    Dim CntcCode As String
    Dim BEDNR As String
    
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
                            oMat01.FlushToDataSource
                            If Trim(oDS_PS_MM010H.GetValue("U_PQType", 0)) = "10" Then
                                oDS_PS_MM010H.setValue "U_RFCAdms", 0, "Y"
                                oDS_PS_MM010H.setValue "U_RFCType", 0, 0
                                For i = 0 To oMat01.VisualRowCount - 2
                                    oDS_PS_MM010L.setValue "U_GuBun", i, 0
                                Next i
                            ElseIf Trim(oDS_PS_MM010H.GetValue("U_PQType", 0)) = "20" Then
                                oDS_PS_MM010H.setValue "U_RFCAdms", 0, "N"
                            End If
                            oMat01.LoadFromDataSource
                        End If
                        
                        '//통합구매
                        If oForm01.Mode = fm_UPDATE_MODE And Trim(oDS_PS_MM010H.GetValue("U_PQType", 0)) = "20" And Trim(oDS_PS_MM010H.GetValue("U_RFCAdms", 0)) = "Y" Then
                            Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("통합구매요청!", oMat01.VisualRowCount - 2 + 1, False)
                            
                            ErrNum = 101
                            oSeq = 1
                            oCount = 1
                            Seq = 1
                            oMat01.FlushToDataSource
                            
                            RequestDate = Trim(oDS_PS_MM010H.GetValue("U_DocDate", 0))
                            DueDate = Trim(oDS_PS_MM010H.GetValue("U_DocDate", 0))
                            BPLID = Trim(oDS_PS_MM010H.GetValue("U_BPLId", 0))
                            
                            '//담당자 성명(전화번호)
                            CntcCode = Trim(oDS_PS_MM010H.GetValue("U_CntcCode", 0))
                            sQry = "select U_FullName + '(' + right(Isnull(U_OffcTel,''),4) + ')' from [@PH_PY001A] where Code = '" & CntcCode & "'"
                            oRecordSet01.DoQuery sQry
                            
                            BEDNR = oRecordSet01.Fields(0).VALUE
                            
                            
                            
                            For i = 0 To oMat01.VisualRowCount - 2
                                If Trim(oDS_PS_MM010L.GetValue("U_E_BANFN", i)) = "" And (Trim(oDS_PS_MM010L.GetValue("U_E_BNFPO", i)) = "" Or Trim(oDS_PS_MM010L.GetValue("U_E_BNFPO", i)) = "00000") Then
                                    ItemCode = Trim(oDS_PS_MM010L.GetValue("U_ItemCode", i))
                                    FileName = Trim(oDS_PS_MM010L.GetValue("U_FILENAME", i))
                                    Dir = Trim(oDS_PS_MM010L.GetValue("U_Dir", i))
                                    ItmBsort = Trim(oDS_PS_MM010L.GetValue("U_ItemGpCd", i))
                                    
                                    
                                    If Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0)) = "60" Or Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0)) = "50" Then '//고정자산품의(60), 상품품의(50) 추가(2016.12.29 송명규, 이영진과장 요청)
                                    
                                        ItemName = Trim(oDS_PS_MM010L.GetValue("U_ItemName", i))
                                        Size = Trim(oDS_PS_MM010L.GetValue("U_OutSize", i))
                                        Weight = Trim(oDS_PS_MM010L.GetValue("U_Weight", i))
                                        Unit = Trim(oDS_PS_MM010L.GetValue("U_OutUnit", i))
                                        ItemType = ""
                                        RequestNo = Trim(oDS_PS_MM010L.GetValue("U_CGNo", i))
                                        
                                    ElseIf Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0)) = "40" Then '외주제작품의(40) 추가(2014.07.02)
                                    
                                        sQry = "Select U_Size, BuyUnitMsr, FrgnName, ItemName, U_ItmMSort From [OITM] Where ItemCode = '" & Left(ItemCode, 11) & "'"
                                        oRecordSet01.DoQuery sQry
                                        
                                        ItemName = Trim(oDS_PS_MM010L.GetValue("U_ItemName", i))
                                        Size = Trim(oDS_PS_MM010L.GetValue("U_OutSize", i))
                                        Weight = Trim(oDS_PS_MM010L.GetValue("U_Weight", i))
                                        Unit = Trim(oDS_PS_MM010L.GetValue("U_OutUnit", i))
                                        ItemType = Trim(oRecordSet01.Fields("U_ItmMSort").VALUE)
                                        RequestNo = Trim(oDS_PS_MM010L.GetValue("U_CGNo", i))
                                        
                                    Else
                                        sQry = "Select U_Size, BuyUnitMsr, FrgnName, ItemName, U_ItmMSort From [OITM] Where ItemCode = '" & ItemCode & "'"
                                        oRecordSet01.DoQuery sQry
                                        If Trim(oDS_PS_MM010L.GetValue("U_OutSize", i)) = "" Then
                                            Size = Trim(oRecordSet01.Fields("U_Size").VALUE)
                                        Else
                                            Size = Trim(oDS_PS_MM010L.GetValue("U_OutSize", i))
                                        End If
                                    
                                        'ItemName = Trim(oRecordSet01.Fields("FrgnName").VALUE) + ":" + Size
                                        '왜 ":"을 중간에 넣고 품목이름과 규격을 합치나... 내역을 전체 저장하면 될 것을...
                                        '이렇게 하니까 통합구매로 전달할 때 ":"을 기준으로 Split을 하는데 당연히 품목이름만 전달 되지...
                                        '내역을 전체 저장하게 수정
                                        '(2012.05.08 송명규 수정)
                                        ItemName = Trim(oRecordSet01.Fields("ItemName").VALUE)
                                        Weight = Trim(oDS_PS_MM010L.GetValue("U_Weight", i))
                                        Unit = Trim(oRecordSet01.Fields("BuyUnitMsr").VALUE)
                                        ItemType = Trim(oRecordSet01.Fields("U_ItmMSort").VALUE) ''"0404" 'Trim(oDS_PS_MM010L.GetValue("U_ItemType", i))
                                        RequestNo = Trim(oDS_PS_MM010L.GetValue("U_CGNo", i))
                                        DueDate = Trim(oDS_PS_MM010L.GetValue("U_DueDate", i))
                                    End If
                                    
'                                    '//청구번호의 청구 요구일/////////// 2013.01.27 N.G.Y 수정
'                                    sQry = "select Convert(Char(8),U_DueDate,112) from [@PS_MM005h] where U_CGNUM = '" & RequestNo & "'"
'                                    oRecordSet01.DoQuery sQry
'                                    DueDate = oRecordSet01.Fields(0).VALUE
'                                    '///////////////////////////////////
                                    RFC_Sender = Trim(PS_RFC_Sender(BPLID, ItmBsort, ItemCode, ItemName, Size, Weight, Unit, RequestDate, DueDate, ItemType, RequestNo, BEDNR, i, oMat01.VisualRowCount - 2, FileName, Dir))
                                    
                                    If RFC_Sender <> "" Then
                                        oDS_PS_MM010L.setValue "U_E_BANFN", i, Trim(Left(RFC_Sender, InStr(RFC_Sender, "/") - 1))
                                        oDS_PS_MM010L.setValue "U_E_BNFPO", i, Trim(Right(RFC_Sender, Len(RFC_Sender) - InStr(RFC_Sender, "/")))
                                        
                                        If Trim(Left(RFC_Sender, InStr(RFC_Sender, "/") - 1)) = "" Or Trim(Right(RFC_Sender, Len(RFC_Sender) - InStr(RFC_Sender, "/"))) = "" Then
                                            Seq = 0
                                        End If
                                        oDS_PS_MM010L.setValue "U_GuBun", i, Seq
                                        Seq = 1
                                    End If
                                    
'                                    oDS_PS_MM010L.setValue "U_E_BANFN", i, Trim(Left(RFC_Sender, InStr(RFC_Sender, "/") - 1))
'                                    oDS_PS_MM010L.setValue "U_E_BNFPO", i, Trim(Right(RFC_Sender, Len(RFC_Sender) - InStr(RFC_Sender, "/")))
'
'                                    If Trim(Left(RFC_Sender, InStr(RFC_Sender, "/") - 1)) = "" Or Trim(Right(RFC_Sender, Len(RFC_Sender) - InStr(RFC_Sender, "/"))) = "" Then
'                                        Seq = 0
'                                    End If
'                                    oDS_PS_MM010L.setValue "U_GuBun", i, Seq
'                                    Seq = 1
                                    
                                    ProgressBar01.VALUE = ProgressBar01.VALUE + 1
                                    ProgressBar01.Text = ProgressBar01.VALUE & "/" & oMat01.VisualRowCount - 2 + 1 & "건 처리중...!"
                                End If
                            Next
                            oMat01.LoadFromDataSource
                            
                            oDS_PS_MM010H.setValue "U_RFCType", 0, Seq
                            
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
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "CGNo" Then
                            If oMat01.Columns("CGNo").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "FILENAME" Then
                            If oMat01.Columns("FILENAME").Cells(pval.Row).Specific.VALUE = "" Then
                                Dim sFile       As String
                                Dim pos As Integer
                                
                                
                                sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.*", "파일선택", "C:\")
                                If sFile = "" Then
                                    Sbo_Application.StatusBar.SetText "파일을 선택해 주세요.", bmt_Short, smt_Error
                                    Exit Sub
                                Else
'                                    If Mid(Right(sFile, 4), 1, 3) = "pdf" Or Mid(Right(sFile, 4), 1, 3) = "hwp" Then
'                                        oForm01.Items("Dir").Specific.VALUE = sFile
                                        oMat01.Columns("Dir").Cells(pval.Row).Specific.VALUE = sFile
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
                                                        oMat01.Columns("FILENAME").Cells(pval.Row).Specific.VALUE = FileName
                                                    End If
                                                Else
                                                    oMat01.Columns("FILENAME").Cells(pval.Row).Specific.VALUE = FileName
                                                End If
                                            Else
                                                oMat01.Columns("FILENAME").Cells(pval.Row).Specific.VALUE = FileName
                                            End If
                                        Else
                                            oMat01.Columns("FILENAME").Cells(pval.Row).Specific.VALUE = FileName
                                        End If
            

'                                    Else
'                                        Sbo_Application.StatusBar.SetText "pdf파일이 아닙니다.", bmt_Short, smt_Error
'                                        Exit Sub
'                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
                If pval.ItemUID = "Mat01" Then
                   If pval.ColUID = "E_BANFN" Then
                        Dim TempForm01 As Variant
                        '통합구매이며 통합구매요청번호가 있을때
                        If oForm01.Items("PQType").Specific.VALUE = "20" And oMat01.Columns("E_BANFN").Cells(pval.Row).Specific.String <> "" Then
                            Set TempForm01 = New PS_FTP
                            Call TempForm01.LoadForm(oMat01.Columns("CGNo").Cells(pval.Row).Specific.String, oMat01.Columns("E_BANFN").Cells(pval.Row).Specific.String, oMat01.Columns("E_BNFPO").Cells(pval.Row).Specific.String)
                            BubbleEvent = False
                        Else
                                    
                        End If
                   End If
                   
                   If pval.ColUID = "CGNo" Then
                        Dim TempForm02 As Variant
                        If oMat01.Columns("CGNo").Cells(pval.Row).Specific.String <> "" Then
                            Set TempForm02 = New PS_MM010_S
                            Call TempForm02.LoadForm(oMat01.Columns("CGNo").Cells(pval.Row).Specific.String)
                            BubbleEvent = False
                        Else
                                    
                        End If
                   End If
                   
                End If
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
                    '//종전값을 넣기위해
                    oPurchase = Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0))
                    oPQType = Trim(oDS_PS_MM010H.GetValue("U_PQType", 0))
                    
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
                ElseIf pval.ItemUID = "Btn01" Then
'                    If HeaderSpaceLineDel = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
                     
                        
                    If Trim(oDS_PS_MM010H.GetValue("U_PQType", 0)) = "10" And (Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0)) <> "30" Or Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0)) <> "40") Then
                        '자체구매 이고 외주제작, 외주가공이 아닐때
                        If Trim(oDS_PS_MM010H.GetValue("U_BPLId", 0)) = "2" Then
                            '기계사업부
                            Call Print_Report01
                            Call Print_Report03
                        Else
                            Call Print_Report03
                        End If
                    ElseIf Trim(oDS_PS_MM010H.GetValue("U_PQType", 0)) = "20" Then
                        '통합구매
                        Call Print_Report02
                    End If
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "Purchase" Or pval.ItemUID = "BPLId" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_MM010L.Clear
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call Add_MatrixRow(0, False)
                    End If
                    
                    If Trim(oForm01.Items("Purchase").Specific.VALUE) = "40" Then
'                        oMat01.Columns("OutItmCd").Editable = True
                        oMat01.Columns("ItemName").Editable = True
                        oMat01.Columns("OutSize").Editable = True
                        oMat01.Columns("OutUnit").Editable = True
                    Else
'                        oMat01.Columns("OutItmCd").Editable = False
                        oMat01.Columns("ItemName").Editable = False
                        oMat01.Columns("OutSize").Editable = False
                        oMat01.Columns("OutUnit").Editable = False
                    End If
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "PQType" Then
                    If oForm01.Mode = fm_UPDATE_MODE Then
                        If oForm01.Items("PQType").Specific.VALUE = "20" Then
                            oForm01.Items("RFCAdms").Specific.Select "N"
                        End If
                    End If
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
                        If pval.ColUID = "CGNo" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "Qty" Then
                            oMat01.FlushToDataSource
                            ItemCode = Trim(oDS_PS_MM010L.GetValue("U_ItemCode", pval.Row - 1))
                            Qty = oDS_PS_MM010L.GetValue("U_Qty", pval.Row - 1)
                            
                            Calculate_Weight = MDC_PS_Common.Calculate_Weight(ItemCode, Qty, Trim(oForm01.Items("BPLId").Specific.VALUE))
                                                                                  
                            oDS_PS_MM010L.setValue "U_Weight", pval.Row - 1, Calculate_Weight
                            oMat01.LoadFromDataSource
                            oMat01.Columns("Weight").Cells(pval.Row).Click ct_Regular
                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                For i = 0 To oMat01.VisualRowCount - 1
                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                Next i
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
                
'                Add_MatrixRow oMat01.RowCount, False
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
                Set oDS_PS_MM010H = Nothing
                Set oDS_PS_MM010L = Nothing
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
                For i = 0 To oMat01.VisualRowCount - 2
                    sQry = "Select b.DocNum From [@PS_MM030L] a Inner Join [@PS_MM030H] b On a.DocEntry = b.DocEntry "
                    sQry = sQry & "where a.U_PQDocNum = '" & Trim(oForm01.Items("DocNum").Specific.VALUE) & "' "
                    sQry = sQry & "And a.U_PQLinNum = '" & Trim(oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE) & "' And b.Canceled = 'N'"
                    oRecordSet01.DoQuery sQry
                    
                    If oRecordSet01.Fields(0).VALUE = 0 Or oRecordSet01.Fields(0).VALUE = "" Then
                    Else
                        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 구매견적이 발주등록 되었습니다. 취소할 수 없습니다.", "E"
                        BubbleEvent = False
                        Exit Sub
                    End If
                Next i
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
                    oDS_PS_MM010L.RemoveRecord oDS_PS_MM010L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                oForm01.Items("SumQty").Specific.VALUE = 0
                oForm01.Items("SumWeight").Specific.VALUE = 0
                
                '//아이디별 사업장 세팅
                oForm01.Items("BPLId").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
                '//아이디별 사번 세팅
                If MDC_PS_Common.User_SuperUserYN = "N" Then '수퍼유저인 경우는 사번 미표기(2016.01.08 송명규)
                    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
                End If
                oForm01.Freeze False
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                oDS_PS_MM010H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                Call Add_MatrixRow(0, True)
                oForm01.Items("SumQty").Specific.VALUE = 0
                oForm01.Items("SumWeight").Specific.VALUE = 0
'                oForm01.Items("BPLId").Click ct_Collapsed
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("CGNo").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_MM010H.GetValue("Status", 0) = "O" Then
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
    Set oDS_PS_MM010H = oForm01.DataSources.DBDataSources("@PS_MM010H")
    Set oDS_PS_MM010L = oForm01.DataSources.DBDataSources("@PS_MM010L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    
    
    oDS_PS_MM010H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
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
    
    Call oForm01.DataSources.UserDataSources.Add("SumQty", dt_SUM)
    oForm01.Items("SumQty").Specific.DataBind.SetBound True, "", "SumQty"
    
    Call oForm01.DataSources.UserDataSources.Add("SumWeight", dt_QUANTITY)
    oForm01.Items("SumWeight").Specific.DataBind.SetBound True, "", "SumWeight"
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//견적형태
    Set oCombo = oForm01.Items("PQType").Specific
    sQry = "SELECT Code, Name From [@PSH_RETYPE]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oDS_PS_MM010H.setValue("U_PQType", 0, "")
    'Call oCombo.Select("0", psk_Index)
    
    Set oCombo = oForm01.Items("RotateYN").Specific
    oCombo.ValidValues.Add "N", "[N]일반품"
    oCombo.ValidValues.Add "Y", "[Y]순환품"
    Call oDS_PS_MM010H.setValue("U_RotateYN", 0, "N")
    
    '//구매방식, 품목구분
    Set oCombo = oForm01.Items("Purchase").Specific
    sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("ItemGpCd").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oDS_PS_MM010H.setValue("U_Purchase", 0, "")
'    oCombo.Select "0", psk_Index
    
    
    
    
    '//품목대분류
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("ItmBSort").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//매입기준단위
    sQry = "SELECT Code, Name From [@PSH_UOMORG] Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("OBasUnit").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//외주사유코드
    sQry = "            SELECT      T1.U_Minor AS [Code],"
    sQry = sQry & "                 T1.U_CdName AS [Value]"
    sQry = sQry & "  FROM       [@PS_SY001H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_SY001L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T0.Code = 'P201'"
    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  T1.U_Seq"
    
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oMat01.Columns("OutCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
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
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
    '//품목구분 종전입력값
    If Trim(oForm01.Items("Purchase").Specific.VALUE) = "" And oPurchase <> "" Then
        Call oForm01.Items("Purchase").Specific.Select(oPurchase, psk_ByValue)
    End If
    
    '//견적형태 종전입력값
    If Trim(oForm01.Items("PQType").Specific.VALUE) = "" And oPQType <> "" Then
        Call oForm01.Items("PQType").Specific.Select(oPQType, psk_ByValue)
    End If
    
    Set oCombo = oForm01.Items("RotateYN").Specific
    oCombo.Select "N", psk_ByValue
    
        
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
        oForm01.Items("PQType").Enabled = True
        oForm01.Items("Purchase").Enabled = True
        oForm01.Items("DocDate").Enabled = True
                
        oForm01.Items("Mat01").Enabled = True
'        oMat01.Columns("CGNo").Editable = True
'        oMat01.Columns("Qty").Editable = True
'        oMat01.Columns("Comments").Editable = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("PQType").Enabled = True
        oForm01.Items("Purchase").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        
        oForm01.Items("Mat01").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        
        If Trim(oDS_PS_MM010H.GetValue("U_PQType", 0)) = "10" Then
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("PQType").Enabled = True
            oForm01.Items("Purchase").Enabled = False
            oForm01.Items("DocDate").Enabled = True
'            oForm01.Items("RFCAdms").Enabled = True
            oForm01.Items("Mat01").Enabled = True
        ElseIf Trim(oDS_PS_MM010H.GetValue("U_PQType", 0)) = "20" Then
            If Trim(oDS_PS_MM010H.GetValue("U_RFCAdms", 0)) = "Y" Then
                oForm01.Items("BPLId").Enabled = False
                oForm01.Items("CntcCode").Enabled = False
                oForm01.Items("PQType").Enabled = False
                oForm01.Items("Purchase").Enabled = False
                oForm01.Items("DocDate").Enabled = False
                oForm01.Items("Mat01").Enabled = False
            Else
                oForm01.Items("BPLId").Enabled = True
                oForm01.Items("CntcCode").Enabled = True
                oForm01.Items("PQType").Enabled = True
                oForm01.Items("Purchase").Enabled = True
                oForm01.Items("DocDate").Enabled = True
                oForm01.Items("Mat01").Enabled = True
            End If
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
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM010'", "")
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
        oDS_PS_MM010L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM010L.offset = oRow
    oDS_PS_MM010L.setValue "U_LineNum", oRow, oRow + 1
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
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_MM010H.GetValue("U_CntcCode", 0)) & "'"
            oRecordSet01.DoQuery sQry
            
            oDS_PS_MM010H.setValue "U_CntcName", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"
            If oCol = "CGNo" Then
                oForm01.Freeze True
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("CGNo").Cells(oRow).Specific.VALUE) <> "" Then
                    oMat01.FlushToDataSource
                    Add_MatrixRow oMat01.RowCount, False
                    oMat01.Columns("CGNo").Cells(oRow).Click ct_Regular
                End If
                
                sQry = "        SELECT  A.U_ItemCode, "
                sQry = sQry & "         A.U_ItemName,"
                sQry = sQry & "         A.U_FrgnName,"
                sQry = sQry & "         A.U_Qty,"
                sQry = sQry & "         A.U_Weight,"
                sQry = sQry & "         A.U_OrdType,"
                sQry = sQry & "         A.U_Duedate,"
                sQry = sQry & "         B.U_ItmBSort,"
                sQry = sQry & "         B.U_ItmMSort,"
                sQry = sQry & "         B.BuyUnitMsr,"
                sQry = sQry & "         A.U_OutItmCd,"
                sQry = sQry & "         A.U_OutSize,"
                sQry = sQry & "         A.U_OutUnit,"
                sQry = sQry & "         A.U_Auto,"
                sQry = sQry & "         A.U_Comments,"
                sQry = sQry & "         A.U_ProcCode,"
                sQry = sQry & "         A.U_ProcName,"
                sQry = sQry & "         A.U_OutCode,"
                sQry = sQry & "         A.U_OutNote"
                sQry = sQry & " FROM    [@PS_MM005H] A"
                sQry = sQry & "         LEFT JOIN "
                sQry = sQry & "         [OITM] B "
                sQry = sQry & "             ON A.U_ItemCode = B.ItemCode"
                sQry = sQry & " WHERE   A.U_CgNum = '" & Trim(oMat01.Columns("CGNo").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                
                If oRecordSet01.RecordCount = 0 Then
                    MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
                    Set oRecordSet01 = Nothing
                    oForm01.Freeze False
                    Exit Sub
                End If
    
                oMat01.Columns("ItemCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
                oMat01.Columns("ItemName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ItemName").VALUE)
'                oMat01.Columns("FrgnName").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_FrgnName").Value)
                oMat01.Columns("Qty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Qty").VALUE)
                oMat01.Columns("Weight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Weight").VALUE)
                oMat01.Columns("ItemGpCd").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_OrdType").VALUE)
                oMat01.Columns("DueDate").Cells(oRow).Specific.VALUE = Format(Trim(oRecordSet01.Fields("U_DueDate").VALUE), "YYYYMMDD")

                oMat01.Columns("ItmBSort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmBSort").VALUE)
                oMat01.Columns("OBasUnit").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("BuyUnitMsr").VALUE)
                
'                oMat01.Columns("OutItmCd").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_OutItmCd").Value)
                oMat01.Columns("OutSize").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutSize").VALUE)
                oMat01.Columns("OutUnit").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutUnit").VALUE)
                oMat01.Columns("Auto").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_Auto").VALUE)
                
                oMat01.Columns("Comments").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_Comments").VALUE)
                
                oMat01.Columns("ProcCode").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ProcCode").VALUE)
                oMat01.Columns("ProcName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_ProcName").VALUE)
                
                If Trim(oRecordSet01.Fields("U_OutCode").VALUE) <> "" Then
                    Call oMat01.Columns("OutCode").Cells(oRow).Specific.Select(Trim(oRecordSet01.Fields("U_OutCode").VALUE)) '외주사유코드
                End If
                oMat01.Columns("OutNote").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_OutNote").VALUE) '외주사유내용
                
'                oMat01.Columns("ItemType").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItemType").Value)
'                oMat01.Columns("ItmMSort").Cells(oRow).Specific.Select Trim(oRecordSet01.Fields("U_ItmMSort").Value)
'                oMat01.Columns("Weight").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_Weight").Value)
'                oMat01.Columns("Weight").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_Weight").Value)
'                oMat01.Columns("Weight").Cells(oRow).Specific.Value = Trim(oRecordSet01.Fields("U_Weight").Value)
                
                oMat01.Columns("CGNo").Cells(oRow).Click ct_Regular
                
                For i = 0 To oMat01.VisualRowCount - 2
                    If oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE = "" Then
                        SumQty = SumQty
                    Else
                        SumQty = SumQty + oMat01.Columns("Qty").Cells(i + 1).Specific.VALUE
                    End If
                    SumWeight = SumWeight + oMat01.Columns("Weight").Cells(i + 1).Specific.VALUE
                Next i
                
                oMat01.AutoResizeColumns
                
                oForm01.Items("SumQty").Specific.VALUE = SumQty
                oForm01.Items("SumWeight").Specific.VALUE = SumWeight
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
        Case oDS_PS_MM010H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM010H.GetValue("U_CntcCode", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM010H.GetValue("U_PQType", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM010H.GetValue("U_Purchase", 0) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM010H.GetValue("U_DocDate", 0) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
    End Select

'    '마감상태 체크_S(2017.11.23 송명규 추가)
'    If MDC_PS_Common.Check_Finish_Status(oForm01.Items("BPLId").Specific.VALUE, oForm01.Items("DocDate").Specific.VALUE) = False Then
'        ErrNum = 6
'        GoTo HeaderSpaceLineDel_Error
'    End If
'    '마감상태 체크_E(2017.11.23 송명규 추가)
    
    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "담당자는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "견적형태는 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "구매방식은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "전기일은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기일을 확인하고, 회계부서로 문의하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error

    Dim i               As Long
    Dim ErrNum          As Integer
    Dim sQry            As String

    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 And oDS_PS_MM010L.GetValue("U_CGNo", 0) = "" Then
        ErrNum = 2
        GoTo MatrixSpaceLineDel_Error
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_MM010L.GetValue("U_CGNo", i) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
'            Case oDS_PS_MM010L.GetValue("U_Qty", i) = "" Or oDS_PS_MM010L.GetValue("U_Qty", i) = 0
'                ErrNum = 4
'                GoTo MatrixSpaceLineDel_Error
            Case oDS_PS_MM010L.GetValue("U_Weight", i) = 0
                ErrNum = 5
                GoTo MatrixSpaceLineDel_Error
            Case oForm01.Items("PQType").Specific.VALUE = "20" And oForm01.Items("Purchase").Specific.VALUE = "20" And Len(Trim(oDS_PS_MM010L.GetValue("U_ItemCode", i))) <> 11
                '//통합구매요청 이고 부재료이면서 11자리가 아니면 에러
                ErrNum = 6
                GoTo MatrixSpaceLineDel_Error
            
            Case PS_MM010_CheckDate(oMat01.Columns("CGNo").Cells(i + 1).Specific.VALUE) = False '구매요청과 일자 체크
            
                ErrNum = 7
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
        MDC_Com.MDC_GF_Message "수량은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "중량은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "통합구매요청은 통합코드로만 가능합니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        Call Sbo_Application.MessageBox(i + 1 & "행 [" & oMat01.Columns("ItemCode").Cells(i + 1).Specific.VALUE & "]의 구매견적일은 구매요청일과 같거나 늦어야합니다. 확인하십시오." & Chr(13) & "해당 견적은 전체가 등록되지 않습니다.", 1)
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
        If Trim(oDS_PS_MM010L.GetValue("U_CGNo", i)) = "" Then
            oDS_PS_MM010L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function PS_RFC_Sender(ByVal BPLID As String, ByVal ItmBsort As String, ByVal ItemCode As String, ByVal ItemName As String, ByVal Size As String, ByVal Qty As Double, ByVal Unit As String, ByVal RequestDate As String, ByVal DueDate As String, ByVal ItemType As String, ByVal RequestNo As String, ByVal BEDNR As String, ByVal i&, ByVal LastRow&, ByVal FileName As String, ByVal Dir As String) As String
On Error GoTo RFC_Sender_Error
    Dim ReturnValue As String
    Dim WERKS As String
    Dim Rotate As String
    
    Dim Comments$
    Dim oFileName  As String
    Dim pos As Integer
    
    If oCount = 1 Then
        Set oSapConnection01 = CreateObject("SAP.Functions")
        oSapConnection01.Connection.User = "ifuser"
        oSapConnection01.Connection.Password = "pdauser"
'        oSapConnection01.Connection.Client = "710"    '테스트
        oSapConnection01.Connection.Client = "210"     '운영
'        oSapConnection01.Connection.ApplicationServer = "192.1.11.7"  '테스트
'        oSapConnection01.Connection.ApplicationServer = "192.1.11.3"   '운영 '디버깅용
        oSapConnection01.Connection.ApplicationServer = "192.1.1.217"  '운영(실제사용)
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
'            oSapConnection01.Connection.Client = "710"  '테스트
            oSapConnection01.Connection.Client = "210" '운영
'            oSapConnection01.Connection.ApplicationServer = "192.1.11.7"    '테스트
'            oSapConnection01.Connection.ApplicationServer = "192.1.11.3"   '운영 '디버깅용
            oSapConnection01.Connection.ApplicationServer = "192.1.1.217"  '운영(실제사용)
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
'                oSapConnection01.Connection.Client = "710"   '테스트
                oSapConnection01.Connection.Client = "210"  '운영
'                oSapConnection01.Connection.ApplicationServer = "192.1.11.7"  '테스트
'               oSapConnection01.Connection.ApplicationServer = "192.1.11.3"   '운영 '디버깅용
                oSapConnection01.Connection.ApplicationServer = "192.1.1.217"  '운영(실제사용)
                oSapConnection01.Connection.language = "KO"
                oSapConnection01.Connection.SystemNumber = "01"
'           oSapConnection01.Connection.SystemNumber = "00" '192.1.11.3 일때는 "00"
                If Not oSapConnection01.Connection.Logon(0, True) Then
                    MDC_Com.MDC_GF_Message "안강(R/3)서버에 접속할수 없습니다.", "E"
                    GoTo RFC_Sender_Exit
                End If
            Else
            End If
        End If
    End If
    
    If oForm01.Items("RotateYN").Specific.VALUE = "Y" Then
        Rotate = "CB"   '//순환품구매요청'
    Else
        Rotate = "NB"   '//표준구매요청
    End If
    
    Dim oFunction01 As Object
    If (ItmBsort = "20" And Len(ItemCode) = 11) Or ItmBsort = "50" Then
        Set oFunction01 = oSapConnection01.Add("ZMM_SUB_PR")  '//부자재이면서 11자리 인경우 통구실 구매요청
    Else
        Set oFunction01 = oSapConnection01.Add("ZMM_INTF_GROUP") '//그외는 자재코드 맵핑해서 처리
    End If
    
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
    
    '비고
    Comments = Trim(oDS_PS_MM010L.GetValue("U_Comments", i))
    
    oFunction01.Exports("I_WERKS") = WERKS '//플랜트 홀딩스 창원 9200, 홀딩스 부산 9300, 포장사업팀 9500, 포장온산 9600
    
    
    oFunction01.Exports("I_MATNR") = ItemCode '//자재코드 char(18)
    
    '통합구매로 본사R3에 값을 전달할 때 품목길이 Check (2012.03.28 송명규)
    If Len(ItemName) <= 40 Then 'ItemName의 길이가 40이하일 경우에만 I_MAKTX 에 전달
    
        oFunction01.Exports("I_MAKTX") = ItemName '//자재내역 char(40)
        oFunction01.Exports("I_WRKST") = Size '//자재규격 char(48)
        
    Else '40 초과일 경우는 I_WRKST에 ItemName에서 ":" 뒤의 규격만 전달
    
        Dim arrItemName() As String
        arrItemName = Split(ItemName, ":")
        
        oFunction01.Exports("I_MAKTX") = arrItemName(0)
        
        If Size = "" Then 'Size에 값이 없을 때만
            oFunction01.Exports("I_WRKST") = arrItemName(1) '//자재규격 char(48)
        Else 'Size에 값이 있을 때는 Size의 값을 전달
            oFunction01.Exports("I_WRKST") = Size '//자재규격 char(48)
        End If
        
    End If
    '통합구매로 본사R3에 값을 전달할 때 품목길이 Check (2012.03.28 송명규)
    
    oFunction01.Exports("I_ZUSE") = Comments '//비고 char(132)"
    oFunction01.Exports("I_MENGE") = Qty '//구매요청수량 dec(13,3)
    oFunction01.Exports("I_MEINS") = Unit '//단위 char(3)
    oFunction01.Exports("I_BADAT") = RequestDate '//구매요청일 char(8)
    oFunction01.Exports("I_LFDAT") = DueDate '//납품일 char(8)
    oFunction01.Exports("I_MATKL") = ItemType '//자재그룹 char(9)
    oFunction01.Exports("I_ZBANFN") = RequestNo '//구매요청번호
    
    oFunction01.Exports("I_BEDNR") = BEDNR '//구매담당자(전화번호)
    
'    oFunction01.Exports("I_FILENAME") = FileName '//FileName 파일명을 보내니까 에러남. 막아도 문제없음. 김동섭과장확인 2015/11/24 노근용
    
    
    oFunction01.Exports("I_BSART") = Rotate '//순환품 구매요청
    
    If Not (oFunction01.Call) Then
        MDC_Com.MDC_GF_Message "안강(R/3)서버 함수호출중 오류발생", "E"
        GoTo RFC_Sender_Exit
    Else
        If (oFunction01.Imports("E_MESSAGE").VALUE = "") Then '//에러메시지
            ReturnValue = oFunction01.Imports("E_BANFN").VALUE & "/" & oFunction01.Imports("E_BNFPO").VALUE '//통합구매요청번호 '//통합구매요청 품목번호
        Else
            oDS_PS_MM010L.setValue "U_MESSAGE", i, "" '초기화

            oDS_PS_MM010L.setValue "U_MESSAGE", i, Trim(oFunction01.Imports("E_MESSAGE").VALUE)
            Call MDC_Com.MDC_GF_Message(oFunction01.Imports("E_MESSAGE").VALUE, "E")
            GoTo RFC_Sender_Exit
        End If
    End If
    
    If Trim(FileName) <> "" Then
        '//FTP파일 전송
        hOpen = 0
        hConnection = 0
        hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
        Dim nFlag As Long
        Dim FDir As String
        
        
        nFlag = INTERNET_FLAG_PASSIVE
        hConnection = InternetConnect(hOpen, FTP_SERVER, INTERNET_INVALID_PORT_NUMBER, _
             FTP_USER, FTP_PASSWORD, INTERNET_SERVICE_FTP, nFlag, 0)
             
        FDir = Dir
        Dim bRet As Boolean
        
        oFileName = FileName
        pos = InStr(oFileName, ".")

        'FileName = WERKS + "_" + ItemCode + Trim(Mid(oFileName, pos, 10))
        FileName = Left(WERKS, 2) + "_" + ItemCode + Trim(Mid(oFileName, pos, 10))
        
        bRet = FtpPutFile(hConnection, FDir, "trans/group/" & FileName, ASCII, 0)
        If bRet = False Then
            '//FTP 파일전송에러
            oDS_PS_MM010L.setValue "U_MESSAGE1", i, "" '초기화

            oDS_PS_MM010L.setValue "U_MESSAGE1", i, "FTP전송오류"
            Call MDC_Com.MDC_GF_Message("FTP전송오류", "E")
            
            'FTP 세션 종료처리
            If hConnection <> 0 Then InternetCloseHandle (hConnection)
            If hOpen <> 0 Then InternetCloseHandle (hOpen)
            hConnection = 0
            hOpen = 0
            'FTP 세션 종료처리 종료
        Else
            '//정상적으로 전송후 처리
        
            'FTP 세션 종료처리
            
            If hConnection <> 0 Then InternetCloseHandle (hConnection)
            If hOpen <> 0 Then InternetCloseHandle (hOpen)
            hConnection = 0
            hOpen = 0
            'FTP 세션 종료처리 종료
        
            Set oFunction01 = oSapConnection01.Add("ZMM_INTF_GROUP_FILE")
    
            If oFunction01 Is Nothing Then
                '//함수호출 에러
                oDS_PS_MM010L.setValue "U_MESSAGE1", i, "" '초기화
        
                oDS_PS_MM010L.setValue "U_MESSAGE1", i, "함수(ZMM_INTF_GROUP_FILE) 생성오류"
                MDC_Com.MDC_GF_Message "함수(ZMM_INTF_GROUP_FILE) 생성오류.", "E"
                
            Else
                '//함수정상호출 하면
                oFunction01.Exports("I_MATNR") = ItemCode  '품목코드
                oFunction01.Exports("I_WERKS") = WERKS     '사업장
                oFunction01.Exports("I_FILENAME") = FileName
                
                If Not oFunction01.Call Then
                    '//함수호출 에러
                    oDS_PS_MM010L.setValue "U_MESSAGE1", i, "" '초기화
        
                    oDS_PS_MM010L.setValue "U_MESSAGE1", i, "통합구매(R/3) 함수(ZMM_INTF_GROUP_FILE)호출중 오류발생"
                    MDC_Com.MDC_GF_Message "통합구매(R/3) 함수(ZMM_INTF_GROUP_FILE)호출중 오류발생", "E"
                    
                Else
                    '//정상메세지 처리
                    oDS_PS_MM010L.setValue "U_MESSAGE1", i, "" '초기화
        
                    oDS_PS_MM010L.setValue "U_MESSAGE1", i, Trim(oFunction01.Imports("E_MESSAGE").VALUE)
                    Call MDC_Com.MDC_GF_Message(oFunction01.Imports("E_MESSAGE").VALUE, "E")
                End If
            End If
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

Private Sub Print_Report01()
On Error GoTo Print_Report01_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim oText(1)        As String
    
    Dim DocNum$, PQType$, Purchase$
    
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
    DocNum = Trim(oDS_PS_MM010H.GetValue("DocNum", 0))
    PQType = Trim(oDS_PS_MM010H.GetValue("U_PQType", 0))
    Purchase = Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0))
    
    '// Crystal
    
    WinTitle = "[PS_MM010]" & "견적의뢰서(결재)"
    ReportName = "PS_MM010_01.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    If Purchase = "10" Then
        oText(1) = "작번 / 품명 / Main 품명"
    ElseIf Purchase = "20" Then
        oText(1) = "사용처"
    End If
    
    For i = 1 To 1
        If Len("" & i & "") = 1 Then
            gRpt_Formula(i) = "F0" & i & ""
        Else
            gRpt_Formula(i) = "F" & i & ""
        End If
        gRpt_Formula_Value(i) = oText(i)
    Next i
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    '// 조회조건문
    sQry = "EXEC PS_MM010_01 '" & DocNum & "', '" & PQType & "'"
    
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Report01_Error
'    End If

    '// Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
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
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub Print_Report02()
On Error GoTo Print_Report02_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim oText(1)        As String
    
    Dim DocNum$, PQType$, Purchase$
    
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
    DocNum = Trim(oDS_PS_MM010H.GetValue("DocNum", 0))
    PQType = Trim(oDS_PS_MM010H.GetValue("U_PQType", 0))
    Purchase = Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0))
    
    '// Crystal
    
    WinTitle = "[PS_MM010]" & "구매요청서(통합구매)"
    ReportName = "PS_MM010_02.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    If Purchase = "10" Then
        oText(1) = "작번 / 품명 / Main 품명"
    ElseIf Purchase = "20" Then
        oText(1) = "사용처"
    End If
    
    For i = 1 To 1
        If Len("" & i & "") = 1 Then
            gRpt_Formula(i) = "F0" & i & ""
        Else
            gRpt_Formula(i) = "F" & i & ""
        End If
        gRpt_Formula_Value(i) = oText(i)
    Next i
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    '// 조회조건문
    sQry = "EXEC PS_MM010_01 '" & DocNum & "', '" & PQType & "'"
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Report02_Error
'    End If

    '// Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report02_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub Print_Report03()
On Error GoTo Print_Report03_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim oText(1)        As String
    
    Dim DocNum$, PQType$, Purchase$
    
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    Call ConnectODBC
    
    DocNum = Trim(oDS_PS_MM010H.GetValue("DocNum", 0))
    PQType = Trim(oDS_PS_MM010H.GetValue("U_PQType", 0))
    Purchase = Trim(oDS_PS_MM010H.GetValue("U_Purchase", 0))
    
    '// Crystal
    
    WinTitle = "[PS_MM010]" & "견적의뢰서"
    ReportName = "PS_MM010_03.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    '// 조회조건문
    sQry = "EXEC PS_MM010_01 '" & DocNum & "', '00'"
    
'    oRecordSet01.DoQuery sQry
'    If oRecordSet01.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Report03_Error
'    End If

    '// Action
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "2", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Report03_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Report03_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Function PS_MM010_CheckDate(ByVal pBaseEntry As String) As Boolean
'******************************************************************************
'Function ID : PS_MM010_CheckDate()
'해당모듈    : PS_MM010
'기능        : 선행프로세스와 일자 비교
'인수        : 없음
'반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_MM010_CheckDate_Error

    Dim Query01      As String
    Dim loopCount    As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BaseEntry    As String
    Dim BaseLine     As String
    Dim DocType      As String
    Dim CurDocDate   As String
    
    BaseEntry = pBaseEntry
    BaseLine = ""
    DocType = "PS_MM010"
    CurDocDate = oForm01.Items("DocDate").Specific.VALUE

    Query01 = "         EXEC PS_Z_CHECK_DATE '"
    Query01 = Query01 & BaseEntry & "','"
    Query01 = Query01 & BaseLine & "','"
    Query01 = Query01 & DocType & "','"
    Query01 = Query01 & CurDocDate & "'"

    Call oRecordSet01.DoQuery(Query01)

    If oRecordSet01.Fields("ReturnValue").VALUE = "False" Then
        PS_MM010_CheckDate = False
    Else
        PS_MM010_CheckDate = True
    End If
    
    Set oRecordSet01 = Nothing

    Exit Function
PS_MM010_CheckDate_Error:
    PS_MM010_CheckDate = False
    Call Sbo_Application.SetStatusBarMessage("PS_MM010_CheckDate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function
