VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReparationValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim missingCallNo As String
Dim theReparationPriceHasValueAndReparationPaymentMethodWithinTheWarranty As String
Dim reparationDateIsMissingOrNull As String
Dim reparationRegisterationDateIsMissingOrNull As String
Dim reparationTimeBeginIsNotValid As String
Dim reparationTimeEndIsNotValid As String
Dim reparationPaymentMethodIsMissing As String
Dim reparationTeamIsMissing As String
Dim reparationWorkStatusIsMissing As String
Dim modelIsNotMatchedWithFamily As String
Dim reparationNotesIsMissingOrEmpty As String
Dim modelItemNoIsMissing As String
Dim modelGasNoIsMissing As String
Dim modelSerialNumberIsMissing As String
Dim modelPurchaseDateIsSmallerThanProductDate As String
Dim modelProductDateIsSmallerThan30Years As String
Dim timeBeginIsBiggerThanTimeEnd As String
Dim modelPurcahseDateIsMissing As String
Dim modelProductDateIsMissing As String
Dim reparationPriceIsLessThanPiecesPrice As String

Enum PayMethodEnum
    Cash = 0
    Credit = 1
    Warranty = 2
End Enum

Private Sub Class_Initialize()
    missingCallNo = "áã íÊã ÅÏÎÇá ÑÞã ÇáÔßæì"
    theReparationPriceHasValueAndReparationPaymentMethodWithinTheWarranty = "áÇ íãßä Ãä Êßæä ÇáÞÓíãÉ Öãä ÇáßÝÇáÉ æ ÇáÞíãÉ ÃßÈÑ ãä ÇáÕÝÑ"
    'reparationDateIsMissingOrNull = "ÊÇÑíÎ ÇáÇÕáÇÍ ÛíÑ ÕÍíÍ"
    reparationRegisterationDateIsMissingOrNull = "ÊÇÑíÎ ÇáÇÏÎÇá ÛíÑ ÕÍíÍ"
    reparationTimeBeginIsNotValid = "ÅÏÎÇá ÓÇÚÉ ÇáÈÏÃ ÛíÑ ÕÍíÍ"
    reparationTimeEndIsNotValid = "ÅÏÎÇá ÓÇÚÉ ÇáÇäÊåÇÁ ÛíÑ ÕÍíÍ"
    reparationPaymentMethodIsMissing = "áã íÊã ÊÍÏíÏ ØÑíÞå ÇáÏÝÚ"
    reparationTeamIsMissing = "áã íÊã ÊÍÏíÏ ÇáæÑÔå"
    reparationWorkStatusIsMissing = "áã íÊã ÊÍÏíÏ ÍÇáå ÇáÇÕáÇÍ"
    modelIsNotMatchedWithFamily = "ÇáãäÊÌ ÛíÑ ãØÇÈÞ ááÚÇÆáå"
    reparationNotesIsMissingOrEmpty = "áã íÊã ÅÏÎÇá ÇáÇÕáÇÍÇÊ"
    modelItemNoIsMissing = "áÇíæÌÏ ÑÞã ãÍÒäí ááãäÊÌ"
    modelGasNoIsMissing = "áã íÊã ÅÏÎÇá äæÚ ÇáÛÇÒ"
    modelSerialNumberIsMissing = "áã íÊã ÅÏÎÇá ÇáÑÞã ÇáÊÓáÓáí ááãÊÌ"
    modelPurchaseDateIsSmallerThanProductDate = "ÊÇÑíÎ ÇáÔÑÇÁ ÇÕÚÑ ãä ÊÇÑíÍ ÇáÇäÊÇÌ"
    modelProductDateIsSmallerThan30Years = "ÊÇÑíÎ ÇáÇäÊÇÌ ÃÕÚÑ ãä ÊÇÑíÎ ÇáÇÕáÇÍ ÈÇßËÑ ãä 30 ÚÇã"
    timeBeginIsBiggerThanTimeEnd = "áÇ íãßä Ãä íßæä æÞÊ ÇáÇäÊåÇÁ ÃÕÛÑ ãä æÞÊ ÇáÈÏÁ"
    modelPurcahseDateIsMissing = "ÊÇÑíÎ ÇáÔÑÇÁ ÛíÑ ÕÍíÍ"
    modelProductDateIsMissing = "ÊÇÑíÎ ÇáÇäÊÇÌ ÛíÑ ÕÍíÍ"
    reparationPriceIsLessThanPiecesPrice = "Þíãå ÇáÇÕáÇÍ ÃÞá ãä ÇÓÚÇÑ ÇáãæÇÏ"
End Sub

Function GetReparationValidResult(reparationInfo As ReparationViewModel) As Collection
Dim validateResult As New Collection
With reparationInfo
    If Val(.CallNo) = 0 Then
     validateResult.Add missingCallNo
    End If
    If Not IsDate(.regestdate) Or IsNull(.regestdate) Then
         validateResult.Add reparationRegisterationDateIsMissingOrNull
    End If
    If Not IsDate(.RepTimeBegin) Or IsNull(.RepTimeBegin) Then
     validateResult.Add reparationTimeBeginIsNotValid
    End If
    If Not IsDate(.RepTimeEnd) Or IsNull(.RepTimeEnd) Then
        validateResult.Add reparationTimeEndIsNotValid
    End If
    If IsNull(.ReparationState) Then
        validateResult.Add reparationWorkStatusIsMissing
    End If
    
    If IsNull(.Cash) Then
        validateResult.Add reparationPaymentMethodIsMissing
    End If
    
    If Not IsNull(.Cash) And .Cash = PayMethodEnum.Warranty And Not IsNull(.RepPrice) And .RepPrice > 0 Then
         validateResult.Add theReparationPriceHasValueAndReparationPaymentMethodWithinTheWarranty
    End If
    
    If .TeamNo = 0 Then
         validateResult.Add reparationTeamIsMissing
    End If
    
    If IsDate(.RepTimeBegin) And _
        Not IsNull(.RepTimeBegin) And _
        IsDate(.RepTimeEnd) And _
        Not IsNull(.RepTimeEnd) And _
        DateDiff("n", .RepTimeBegin, .RepTimeEnd) < 0 Then
        validateResult.Add timeBeginIsBiggerThanTimeEnd
    End If
    If .Cash <> PayMethodEnum.Warranty Then
        Dim sumAllItemsPrice As Double
        sumAllItemsPrice = 0
        For i = 1 To .ReparationPiecesViewModel.Count
                 sumAllItemsPrice = sumAllItemsPrice + (.ReparationPiecesViewModel(i).Price * .ReparationPiecesViewModel(i).Qty)
        Next i
        If .RepPrice < sumAllItemsPrice Then
             validateResult.Add reparationPriceIsLessThanPiecesPrice
        End If
    End If
    
End With

Set GetReparationValidResult = validateResult
End Function


