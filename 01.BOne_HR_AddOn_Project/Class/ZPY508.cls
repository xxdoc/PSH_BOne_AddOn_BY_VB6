VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY508"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY508.cls
'//  Module         : ÀÎ»ç°ü¸®>Á¤»ê°ü¸®
'//  Desc           : ¿¬±ÝÀúÃà ¼Òµæ°øÁ¦ ¸í¼¼ µî·Ï
'//  FormType       : 2000060508
'//  Create Date    : 2011.01.03
'//  Modified Date  :
'//  Creator        : Choi Dong Kwon
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************

Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oDS_ZPY508H     As SAPbouiCOM.DBDataSource    '½Ã½ºÅÛÄÚµå Çì´õ
Private oDS_ZPY508L     As SAPbouiCOM.DBDataSource    '½Ã½ºÅÛÄÚµå ¶óÀÎ
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     'Å¬·¡½º¿¡¼­ ¼±ÅÃÇÑ ¸¶Áö¸· ¾ÆÀÌÅÛ Uid°ª
Private Col_Last_Uid    As String                     '¸¶Áö¸·¾ÆÀÌÅÛÀÌ ¸ÞÆ®¸¯½ºÀÏ°æ¿ì¿¡ ¸¶Áö¸· ¼±ÅÃµÈ ColÀÇ Uid°ª
Private Col_Last_Row    As Long                       '¸¶Áö¸·¾ÆÀÌÅÛÀÌ ¸ÞÆ®¸¯½ºÀÏ°æ¿ì¿¡ ¸¶Áö¸· ¼±ÅÃµÈ Row°ª
Private oOLDCHK         As String

'*******************************************************************
' .srf ÆÄÀÏ·ÎºÎÅÍ ÆûÀ» ·ÎµåÇÑ´Ù
'*******************************************************************
Public Sub LoadForm(Optional JSNYER$, Optional MSTCOD$, Optional CLTCOD$)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY508.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    '//¿©·¯°³ÀÇ ¸ÞÆ®¸¯½º°¡ Æ²°æ¿ì¿¡ Ãþ°è¸ð¾çÃ³·³ ·Îµå µÇµµ·Ï ¸¸µç ¸ð¾ç
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY508_" & GetTotalFormsCount
    
    'Æû ÇÒ´ç
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    'ÄÃ·º¼Ç¿¡ ÆûÀ» ´ã´Â´Ù   **ÄÃ·º¼ÇÀÌ¶õ °³Ã¼¸¦ ´ã¾Æ ³õ´Â ¹è¿­·Î¼­ ¿©±â¼­´Â È°¼ºÈ­µÇ¾îÁ® ÀÖ´Â ÆûÀ» ´ã°í ÀÖ´Ù
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    AddForms Me, oFormUniqueID, "ZPY508"
    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    'È­¸éÅ°°ª(È­¸é¿¡¼­ À¯ÀÏÅ°°ªÀ» ´ã°í ÀÖ´Â ¾ÆÀÌÅÛÀÇ Uid°ª)
    oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    CreateItems
    
    oForm.EnableMenu ("1293"), True  '/ Çà»èÁ¦
    oForm.EnableMenu ("1283"), True  '/ Á¦°Å
    oForm.EnableMenu ("1284"), False  '/ Ãë¼Ò
    
    
    If JSNYER <> "" Then
        Call ShowSource(JSNYER, MSTCOD, CLTCOD)
    End If
    
    oForm.Freeze False
    oForm.Update
    'oForm.Visible = True

    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
    Sbo_Application.StatusBar.SetText "Form_Load Error:" & Err.Description, bmt_Short, smt_Error
    If (oForm Is Nothing) = False Then
        oForm.Freeze False
        Set oForm = Nothing
    End If
End Sub
'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
'et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_ITEM_PRESSED
            If pval.BeforeAction Then
                If pval.ItemUID = "1" Then
                        '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
                        '//Ãß°¡¹× ¾÷µ¥ÀÌ½Ã¿¡
                        '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
                        If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            If oForm.Mode = fm_ADD_MODE Then
                                If MDC_SetMod.Value_ChkYn("[@ZPY508H]", "U_JSNYER", "'" & oForm.Items("JSNYER").Specific.String & "'", " AND U_MSTCOD = '" & oForm.Items("MSTCOD").Specific.String & "'") = False Then
                                    Sbo_Application.StatusBar.SetText "ÀÌ¹Ì ÀúÀåµÇ¾îÁ® ÀÖ´Â Çì´õÀÇ ³»¿ë°ú ÀÏÄ¡ÇÕ´Ï´Ù.", bmt_Short, smt_Error
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        If Trim$(oDS_ZPY508H.GetValue("U_ENDCHK", 0)) = "Y" And Trim$(oOLDCHK) = "Y" Then
                            Sbo_Application.StatusBar.SetText "Àá±Ý ÀÚ·áÀÔ´Ï´Ù. È®ÀÎÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
                             BubbleEvent = False
                             Exit Sub
                        ElseIf MatrixSpaceLineDel = False Then
                                BubbleEvent = False
                            End If
                        End If
                ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then '/ ChooseBtn»ç¿ø¸®½ºÆ®
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            Else
                If pval.ItemUID = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
                    oForm.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                End If
            End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
            If pval.BeforeAction = True And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" And oForm.Mode <> fm_FIND_MODE Then
                If Last_Item = "MSTCOD" Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(Last_Item).Specific.String & "'", "") = True _
                    And oForm.Items(Last_Item).Specific.String <> "" And Last_Item <> pval.ItemUID Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            If pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And Last_Item = "Mat1" And Col_Last_Uid = "Col1" And Col_Last_Row > 0 _
              And (Col_Last_Uid <> pval.ColUID Or Col_Last_Row <> pval.Row) And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" Then
                If Col_Last_Row > oMat1.VisualRowCount Then
                    Exit Sub
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If (pval.ItemUID = "MSTCOD" Or pval.ItemUID = "JSNYER") Then
                    FlushToItemValue pval.ItemUID
                ElseIf pval.ItemUID = "Mat1" And (pval.ColUID = "Col5") Then
                    FlushToItemValue pval.ColUID, pval.Row
                End If
            End If
        
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
            '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
            '//Ãß°¡¸ðµå¿¡¼­ ÄÚµåÀÌº¥Æ®°¡ ÄÚµå¿¡¼­ ÀÏ¾î ³µÀ»¶§
            '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
             If pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
                If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(pval.ItemUID).Specific.String & "'", "") = True Then
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                Else
                    If oMat1.RowCount > 0 Then
                        oMat1.Columns("Col5").Cells(oMat1.VisualRowCount).CLICK ct_Regular
                        BubbleEvent = False
                    End If
                End If
             End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUID
                    Col_Last_Row = pval.Row
                    Col_Last_Uid = pval.ColUID
                End If
            Else
                Last_Item = pval.ItemUID
                Col_Last_Row = 0
                Col_Last_Uid = ""
            End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
            'ÄÃ·º¼Ç¿¡¼­ »èÁ¦¹× ¸ðµç ¸Þ¸ð¸® Á¦°Å
            '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_ZPY508H = Nothing
                Set oDS_ZPY508L = Nothing
                Set oMat1 = Nothing
            End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_MATRIX_LOAD:
            If pval.BeforeAction = False Then
                FormItemEnabled
                Call Matrix_AddRow(oMat1.VisualRowCount)
            End If
            
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
End Sub
'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i&
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
            Case "1283" '/ Á¦°Å
                If Trim$(oDS_ZPY508H.GetValue("U_ENDCHK", 0)) = "Y" Then
                    Sbo_Application.StatusBar.SetText "Àá±Ý ÀÚ·áÀÔ´Ï´Ù. È®ÀÎÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
                    BubbleEvent = False
                    Exit Sub
                Else
                    If Sbo_Application.MessageBox("ÇöÀç È­¸é³»¿ëÀüÃ¼¸¦ Á¦°Å ÇÏ½Ã°Ú½À´Ï±î? º¹±¸ÇÒ ¼ö ¾ø½À´Ï´Ù.", 2, "Yes", "No") = 2 Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Case "1288", "1289", "1290", "1291"
                Call AuthorityCheck(oForm, "CLTCOD", "@ZPY508H", "DocNum")      '//Á¢¼ÓÀÚ ±ÇÇÑ¿¡ µû¸¥ »ç¾÷Àå º¸±â
                
            
            Case Else
                Exit Sub
        End Select
    Else
    
        Select Case pval.MenuUID
            Case "1287" '/ º¹Á¦
            Case "1283" '/ Á¦°Å
                 FormItemEnabled
            Case "1281", "1282":
                FormItemEnabled
                If pval.MenuUID = "1282" Then
                    FormClear
                    Call Matrix_AddRow(0, True)
                    oForm.Items("JSNYER").CLICK ct_Regular
                End If
            Case "1288" To "1291":
            Case "1293"
                If oMat1.RowCount <> oMat1.VisualRowCount Then
                    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
                    '//¸Ç¸¶Áö¸·¿¡ µ¥ÀÌÅÍ¸¦ »èÁ¦ÇÏ´Â ÀÌÀ¯´Â ÇàÀ» Ãß°¡ ÇÒ°æ¿ì¿¡ µðºñµ¥ÀÌÅÍ¼Ò½º¿¡
                    '//ÀÌ¹Ì µ¥ÀÌÅÍ°¡ µé¾î°¡ ÀÖ±â ¶§¹®¿¡ ÀúÀå½Ã¿¡´Â ¸¶Áö¸· Çà(DBµ¥ÀÌÅÍ ¼Ò½º¿¡)À» »èÁ¦ÇÑ´Ù
                    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
                    For i = 0 To oMat1.VisualRowCount - 1
                        oMat1.Columns("Col0").Cells(i + 1).Specific.Value = i + 1
                    Next i
                    
                    oMat1.FlushToDataSource
                    oDS_ZPY508L.RemoveRecord oDS_ZPY508L.Size - 1   '// Mat1¿¡ ¸¶Áö¸·¶óÀÎ(ºó¶óÀÎ) »èÁ¦
                    oMat1.Clear
                    oMat1.LoadFromDataSource
                    
                End If
                Call FlushToItemValue("Col3", 1)
        End Select
    End If
End Sub


Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim i       As Long
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select

    End If
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
Raise_FormDataEvent_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub
'*******************************************************************
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oColumn         As SAPbouiCOM.Column
    Dim oCheck          As SAPbouiCOM.CheckBox
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//µðºñµ¥ÀÌÅÍ ¼Ò½º °³Ã¼ ÇÒ´ç
    Set oDS_ZPY508H = oForm.DataSources.DBDataSources("@ZPY508H")
    Set oDS_ZPY508L = oForm.DataSources.DBDataSources("@ZPY508L")
    
    Set oMat1 = oForm.Items("Mat1").Specific
    
    '//»ç¾÷Àå
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "ÀüÃ¼"
    oForm.Items("CLTCOD").DisplayDesc = True
    oCombo.Select "%", psk_ByValue
    
    '// °ü°èÄÚµå
    Set oColumn = oMat1.Columns("Col1")
    oColumn.ValidValues.Add "11", "ÅðÁ÷¿¬±Ý-±Ù·ÎÀÚÅðÁ÷±Þ¿©º¸Àå¹ý"
    oColumn.ValidValues.Add "12", "ÅðÁ÷¿¬±Ý-°úÇÐ±â¼úÀÎ°øÁ¦È¸"
    oColumn.ValidValues.Add "21", "¿¬±ÝÀúÃà-°³ÀÎ¿¬±ÝÀúÃà"
    oColumn.ValidValues.Add "22", "¿¬±ÝÀúÃà-¿¬±ÝÀúÃà"
    oColumn.ValidValues.Add "31", "ÁÖÅÃ¸¶·Ã-Ã»¾àÀúÃà"
    oColumn.ValidValues.Add "32", "ÁÖÅÃ¸¶·Ã-ÁÖÅÃÃ»¾àÁ¾ÇÕÀúÃà"
    oColumn.ValidValues.Add "33", "ÁÖÅÃ¸¶·Ã-Àå±âÁÖÅÃ¸¶·ÃÀúÃà"
    oColumn.ValidValues.Add "34", "ÁÖÅÃ¸¶·Ã-±Ù·ÎÀÚÁÖÅÃ¸¶·ÃÀúÃà"
    oColumn.ValidValues.Add "41", "Àå±âÁÖ½ÄÇüÀúÃà"
 
    '// ±ÝÀ¶±â°ü
    Set oColumn = oMat1.Columns("Col2")
    sQry = "SELECT BankCode, BankName FROM [ODSC]"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oColumn.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
        oRecordSet.MoveNext
    Loop
 
    '// ³³ÀÔ¿¬Â÷
    Set oColumn = oMat1.Columns("Col4")
    oColumn.ValidValues.Add "00", "ÇØ´ç¾øÀ½"
    oColumn.ValidValues.Add "01", "1³âÂ÷"
    oColumn.ValidValues.Add "02", "2³âÂ÷"
    oColumn.ValidValues.Add "03", "3³âÂ÷"
    
    '// Á¾(Àü) ¿©ºÎ
    Set oColumn = oMat1.Columns("Col7")
    oColumn.ValOff = "N":   oColumn.ValOn = "Y"
 
   '/ Check ¹öÆ°
    Set oCheck = oForm.Items("ENDCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"

    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oCheck = Nothing
    Set oColumn = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim iRow        As Long
    Dim oMast       As ZPAY_g_EmpID
    Dim TOTAMT      As Double
    
    Select Case oUID
      Case "JSNYER"
            If Trim$(oForm.Items(oUID).Specific.String) <> "" Then
                ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
            Else
                oDS_ZPY508H.setValue "U_JSNYER", 0, ZPAY_GBL_JSNYER
            End If
            oForm.Items(oUID).Update
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oDS_ZPY508H.setValue "U_MSTCOD", 0, ""
                oDS_ZPY508H.setValue "U_MSTNAM", 0, ""
                oDS_ZPY508H.setValue "U_EmpID", 0, ""
                oDS_ZPY508H.setValue "U_CLTCOD", 0, ""
            Else
                oDS_ZPY508H.setValue "U_MSTCOD", 0, UCase$(oForm.Items(oUID).Specific.String)
                oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oDS_ZPY508H.setValue "U_MSTNAM", 0, oMast.MSTNAM
                oDS_ZPY508H.setValue "U_EmpID", 0, oMast.EmpID
                oDS_ZPY508H.setValue "U_CLTCOD", 0, oMast.CLTCOD
            End If
            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
            oForm.Items("CLTCOD").Update
            oForm.Items(oUID).Update
    End Select
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    Select Case oUID
        Case "Col5"
            oMat1.FlushToDataSource

            For iRow = 1 To oMat1.VisualRowCount
                TOTAMT = TOTAMT + Val(oDS_ZPY508L.GetValue("U_SAVAMT", iRow - 1))
            Next iRow
            oDS_ZPY508H.setValue "U_TOTAMT", 0, TOTAMT
            oForm.Items("TOTAMT").Update
    End Select
    If Left$(oUID, 3) = "Col" Then
        oMat1.FlushToDataSource
        If oRow = oMat1.RowCount And Val(oDS_ZPY508L.GetValue("U_SAVAMT", oRow - 1)) <> 0 Then
            Matrix_AddRow oRow
            oMat1.Columns("Col5").Cells(oRow).CLICK ct_Regular
        End If
    End If
End Sub
Private Sub FormClear()
    Dim DocNum&
    
    DocNum = MDC_SetMod.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'ZPY508'", "")
    
    If DocNum = 0 Then
        oForm.Items("DocNum").Specific.String = 1
    Else
        oForm.Items("DocNum").Specific.String = DocNum
    End If
    Call FlushToItemValue("JSNYER")
    
End Sub

Private Sub FormItemEnabled()
    If oForm.Mode = fm_FIND_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = True
        oForm.Items("DocNum").Enabled = True
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_ADD_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_OK_MODE Then
        oForm.Items("JSNYER").Enabled = False
        oForm.Items("MSTCOD").Enabled = False
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        '// ³â¸¶°¨µÈ°ÍÀº ºñÈ°¼ºÈ­
        oOLDCHK = oDS_ZPY508H.GetValue("U_ENDCHK", 0)
        If MDC_SetMod.Get_ReData("U_ENDCHK", "U_JOBYER", "[@ZPY509L]", "'" & oDS_ZPY508H.GetValue("U_JSNYER", 0) & "'", " AND Code = '" & oDS_ZPY508H.GetValue("U_CLTCOD", 0) & "'") = "Y" Then
            oForm.Items("ENDCHK").Enabled = False
        Else
            oForm.Items("ENDCHK").Enabled = True
        End If
        
    End If
End Sub

Private Function MatrixSpaceLineDel() As Boolean
'¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
'ÀúÀåÇÒ µ¥ÀÌÅÍÀÇ À¯È¿¼ºÀ» Á¡°ËÇÑ´Ù
'¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
On Error GoTo Error_Message
    Dim iRow        As Long
    Dim kRow        As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    
    ErrNum = 0
    '/ Çì´õºÎºÐ Ã¼Å©
    Select Case True
      Case Len(Trim$(oDS_ZPY508H.GetValue("U_JSNYER", 0))) <> 4
           ErrNum = 2
           GoTo Error_Message
      Case oDS_ZPY508H.GetValue("U_MSTCOD", 0) = ""
           ErrNum = 3
           GoTo Error_Message
      Case oDS_ZPY508H.GetValue("U_CLTCOD", 0) = ""
           ErrNum = 4
           GoTo Error_Message
           
    End Select
    
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    'È­¸é»óÀÇ ¸ÞÆ®¸¯½º¿¡ ÀÔ·ÂµÈ ³»¿ëÀ» ¸ðµÎ µðºñµ¥ÀÌÅÍ¼Ò½º·Î ³Ñ±ä´Ù
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    oMat1.FlushToDataSource
        
    If oMat1.RowCount = 1 Then          '// Mat1¿¡ °ªÀÌ ÀÖ´ÂÁö È®ÀÎ (ErrorNumber : 1)
        ErrNum = 1
        GoTo Error_Message
    End If
    
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    '//¸¶Áö¸· Çà ÇÏ³ª¸¦ »©°í i=0ºÎÅÍ ½ÃÀÛÇÏ¹Ç·Î ÇÏ³ª¸¦ »©¹Ç·Î
    '//oMat1.RowCount - 2°¡ µÈ´Ù..¹Ýµå½Ã µé¾î °¡¾ß ÇÏ´Â ÇÊ¼ö°ªÀ» È®ÀÎÇÑ´Ù
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    For iRow = 0 To oMat1.VisualRowCount - 2      '// Mat1¿¡ ÀÔ·Â°ªÀÌ ¿Ã¹Ù¸£°Ô µé¾î°¬´ÂÁö È®ÀÎ (ErrorNumber : 3)
        oDS_ZPY508L.Offset = iRow
        If Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) = "" Then
            ErrNum = 6
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY508L.GetValue("U_SAVCOD", iRow)) = "" Then
            ErrNum = 7
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY508L.GetValue("U_SAVNUM", iRow)) = "" Then
            ErrNum = 8
            oMat1.Columns("Col3").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) = "41" And Trim$(oDS_ZPY508L.GetValue("U_STYEAR", iRow)) = "00" Then
            ErrNum = 9
            oMat1.Columns("Col3").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        Else
            '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
            'Áßº¹Ã¼Å©ÀÛ¾÷
            '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
            Chk_Data = Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) & Trim$(oDS_ZPY508L.GetValue("U_SAVCOD", iRow)) & _
                       Trim$(oDS_ZPY508L.GetValue("U_SAVNUM", iRow)) & Trim$(oDS_ZPY508L.GetValue("U_STYEAR", iRow))
            For kRow = iRow + 1 To oMat1.VisualRowCount - 2
                oDS_ZPY508L.Offset = kRow
                If Trim$(Chk_Data) = Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", kRow)) & Trim$(oDS_ZPY508L.GetValue("U_SAVCOD", kRow)) & _
                                    Trim$(oDS_ZPY508L.GetValue("U_SAVNUM", kRow)) & Trim$(oDS_ZPY508L.GetValue("U_STYEAR", kRow)) Then
                    ErrNum = 5
                    oMat1.Columns("Col3").Cells(iRow + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            Next kRow
        End If
        
        If Trim$(oDS_ZPY508L.GetValue("U_SAVGBN", iRow)) <> "41" And Trim$(oDS_ZPY508L.GetValue("U_STYEAR", iRow)) <> "00" Then
            oDS_ZPY508L.Offset = iRow
            oDS_ZPY508L.setValue "U_STYEAR", iRow, "00"
            oMat1.SetLineData (iRow + 1)
        End If
    Next iRow
    
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    '//¸Ç¸¶Áö¸·¿¡ µ¥ÀÌÅÍ¸¦ »èÁ¦ÇÏ´Â ÀÌÀ¯´Â ÇàÀ» Ãß°¡ ÇÒ°æ¿ì¿¡ µðºñµ¥ÀÌÅÍ¼Ò½º¿¡
    '//ÀÌ¹Ì µ¥ÀÌÅÍ°¡ µé¾î°¡ ÀÖ±â ¶§¹®¿¡ ÀúÀå½Ã¿¡´Â ¸¶Áö¸· Çà(DBµ¥ÀÌÅÍ ¼Ò½º¿¡)À» »èÁ¦ÇÑ´Ù
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    oDS_ZPY508L.RemoveRecord oDS_ZPY508L.Size - 1   '// Mat1¿¡ ¸¶Áö¸·¶óÀÎ(ºó¶óÀÎ) »èÁ¦
    
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    'ÇàÀ» »èÁ¦ÇÏ¿´À¸´Ï DBµ¥ÀÌÅÍ ¼Ò½º¸¦ ´Ù½Ã °¡Á®¿Â´Ù
    '¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ¤Ñ
    oMat1.LoadFromDataSource
    
    MatrixSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Select Case ErrNum
    Case 1:     Sbo_Application.StatusBar.SetText "ÀÔ·ÂÇÒ µ¥ÀÌÅÍ°¡ ¾ø½À´Ï´Ù. ÀÔ·ÂÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
    Case 2:     Sbo_Application.StatusBar.SetText "±Í¼Ó³âµµ¸¦ È®ÀÎÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
    Case 3:     Sbo_Application.StatusBar.SetText "ÀÇ·áºñ Áö±Þ±Ý¾×ÀÌ 0ÀÔ´Ï´Ù. ÀÔ·ÂÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
    Case 4:     Sbo_Application.StatusBar.SetText "ÀÚ»çÄÚµå´Â ÇÊ¼öÀÔ´Ï´Ù. ÀÔ·ÂÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
    Case 5:     Sbo_Application.StatusBar.SetText "Áßº¹ÀÔ·ÂµÇ¾ú½À´Ï´Ù. ÀúÃà±¸ºÐ/±ÝÀ¶±â°ü/°èÁÂ¹øÈ£/³³ÀÔ¿¬Â÷º°·Î Áý°èÇÏ¿© ÀÔ·ÂÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
    Case 6:     Sbo_Application.StatusBar.SetText "ÀúÃà±¸ºÐÀº ÇÊ¼öÀÔ´Ï´Ù. ¼±ÅÃÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
    Case 7:     Sbo_Application.StatusBar.SetText "±ÝÀ¶±â°üÀº ÇÊ¼öÀÔ´Ï´Ù. ¼±ÅÃÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
    Case 8:     Sbo_Application.StatusBar.SetText "°èÁÂ¹øÈ£´Â ÇÊ¼öÀÔ´Ï´Ù. ÀÔ·ÂÇÏ¿© ÁÖ½Ê½Ã¿À.", bmt_Short, smt_Error
    Case 9:     Sbo_Application.StatusBar.SetText "Àå±âÁÖ½ÄÇü ÀúÃàÀÎ °æ¿ì ³³ÀÔ¿¬Â÷¸¦ 1³âÂ÷~3³âÂ÷·Î ¼±ÅÃÇÏ¿©¾ß ÇÕ´Ï´Ù.", bmt_Short, smt_Error
    Case Else
        Sbo_Application.StatusBar.SetText "MatrixSpaceLineDel Error:" & Err.Number & Space$(10) & Err.Description, bmt_Short, smt_Error
    End Select
    MatrixSpaceLineDel = False
End Function

Private Sub ShowSource(JSNYER$, MSTCOD$, CLTCOD$)
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim DocEntry$
    Dim oMast           As ZPAY_g_EmpID
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    sQry = "SELECT DocEntry FROM [@ZPY508H]"
    sQry = sQry & "   WHERE U_JSNYER = N'" & JSNYER & "'"
    sQry = sQry & "   AND   U_MSTCOD = N'" & MSTCOD & "'"
    sQry = sQry & "   AND   U_CLTCOD = N'" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            DocEntry = oRecordSet.Fields(0).Value
            oRecordSet.MoveNext
        Loop
        oForm.Items("JSNYER").Specific.Value = JSNYER
        oForm.Items("MSTCOD").Specific.String = MSTCOD
        
        Call oForm.Items("CLTCOD").Specific.Select(CLTCOD, psk_ByValue)
        oForm.Items("DocNum").Specific.Value = DocEntry
        
        oForm.Items("DocNum").Update
        oMat1.LoadFromDataSource
        oForm.Update
        oForm.Items("1").CLICK ct_Regular
        
    Else
        oForm.Mode = fm_OK_MODE
        Call Sbo_Application.ActivateMenuItem("1282")
        
        Call oDS_ZPY508H.setValue("U_JSNYER", 0, JSNYER)
        Call oDS_ZPY508H.setValue("U_MSTCOD", 0, MSTCOD)
        Call oDS_ZPY508H.setValue("U_CLTCOD", 0, CLTCOD)
        oMast = MDC_SetMod.Get_EmpID_InFo(MSTCOD)
        Call oDS_ZPY508H.setValue("U_MSTNAM", 0, oMast.MSTNAM)
        Call oDS_ZPY508H.setValue("U_EmpID", 0, oMast.EmpID)
        
        oForm.Update
        
        Sbo_Application.SendKeys "{TAB}"
    End If
    
    Set oRecordSet = Nothing
    
End Sub

'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oRow As Long, Optional Insert_YN As Boolean)
    If Insert_YN = False Then
        oDS_ZPY508L.InsertRecord (oRow)
    End If
    oDS_ZPY508L.Offset = oRow
    oDS_ZPY508L.setValue "U_LineNum", oRow, oRow + 1
    oDS_ZPY508L.setValue "U_SAVGBN", oRow, ""
    oDS_ZPY508L.setValue "U_SAVCOD", oRow, ""
    oDS_ZPY508L.setValue "U_JONGYN", oRow, "N"
    oDS_ZPY508L.setValue "U_SAVNAM", oRow, ""
    oDS_ZPY508L.setValue "U_SAVNUM", oRow, ""
    oDS_ZPY508L.setValue "U_STYEAR", oRow, "00"
    oDS_ZPY508L.setValue "U_SAVAMT", oRow, 0
    oDS_ZPY508L.setValue "U_SARAMT", oRow, 0
    oMat1.LoadFromDataSource
End Sub


