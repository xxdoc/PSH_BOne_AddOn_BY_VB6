Attribute VB_Name = "SubMain"
Public Application As ZZMDC

Public Sub Main()

    Set Application = New ZZMDC

    Do While GetMessage(Msg, 0&, 0&, 0&)
        TranslateMessage Msg
        DispatchMessage Msg
    DoEvents
    Loop
End Sub

'//����ü�߰�
Public Sub AddForms(ByVal cObject As Variant, ByVal oFormUid As String, Optional ByVal oFormTypeEx)
    MDC_Globals.ClassList.Add cObject, oFormUid
    MDC_Globals.FormTotalCount = MDC_Globals.FormTotalCount + 1
    MDC_Globals.FormCurrentCount = MDC_Globals.FormCurrentCount + 1
    MDC_Globals.FormTypeList.Add oFormTypeEx, Str(MDC_Globals.FormTypeListCount)
    MDC_Globals.FormTypeListCount = MDC_Globals.FormTypeListCount + 1
End Sub

'//����ü����
Public Sub RemoveForms(ByVal oFormUniqueID As String)
    Dim oTempClass          As Variant
    Set oTempClass = MDC_Globals.ClassList.Item(oFormUniqueID)
    MDC_Globals.ClassList.Remove oFormUniqueID
    MDC_Globals.FormCurrentCount = MDC_Globals.FormCurrentCount - 1
    MDC_Globals.FormTypeList.Remove Str(MDC_Globals.FormTypeListCount - 1)
    MDC_Globals.FormTypeListCount = MDC_Globals.FormTypeListCount - 1
    Set oTempClass = Nothing
End Sub

'//�����簴ü��
Public Function GetCurrentFormsCount() As Long
    GetCurrentFormsCount = FormCurrentCount
End Function


'//���Ѱ�ü��
Public Function GetTotalFormsCount() As Long
    GetTotalFormsCount = FormTotalCount
End Function
