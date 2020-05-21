Sub influencer()

'##Dim Social media As String
'##Dim Type of post As String
'##Dim Influencer rate As String

'###############Declarar valores de las hojas
Set ws1 = ActiveWorkbook.Sheets("Digital")


'##################LLenar la Columna de Valor
Row = 2
Col = "O"
Range(Col & Row).Activate

While ActiveCell.value <> 0

    If InStr(1, ActiveCell.value, "InstagramPostNano") Then
        Range("P" & Row).value = "100"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
        '##############################
    ElseIf InStr(1, ActiveCell.value, "InstagramPostMicro") Then
        Range("P" & Row).value = "172"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
        '##################################
    ElseIf InStr(1, ActiveCell.value, "InstagramPostMega") Then
        Range("P" & Row).value = "296"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramPostPower") Then
        Range("P" & Row).value = "507"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramPostSuper Power") Then
        Range("P" & Row).value = "1346"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramPostTop Influencer") Then
        Range("P" & Row).value = "2085"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramPostCelebrity") Then
        Range("P" & Row).value = "2383"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramPostTop Celebrity") Then
        Range("P" & Row).value = "3660"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
'#########################
    ElseIf InStr(1, ActiveCell.value, "InstagramVideoNano") Then
        Range("P" & Row).value = "114"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramVideoMicro") Then
        Range("P" & Row).value = "219"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramVideoMega") Then
        Range("P" & Row).value = "310"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
   ElseIf InStr(1, ActiveCell.value, "InstagramVideoPower") Then
        Range("P" & Row).value = "775"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
   ElseIf InStr(1, ActiveCell.value, "InstagramVideoSuper Power") Then
        Range("P" & Row).value = "1614"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramVideoTop Influencer") Then
        Range("P" & Row).value = "2318"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramVideoCelebrity") Then
        Range("P" & Row).value = "2750"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramVideoTop Celebrity") Then
        Range("P" & Row).value = "4207"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramStoryNano") Then
        Range("P" & Row).value = "43"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramStoryMicro") Then
        Range("P" & Row).value = "73"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramStoryMega") Then
        Range("P" & Row).value = "115"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramStoryPower") Then
        Range("P" & Row).value = "210"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramStorySuper Power") Then
        Range("P" & Row).value = "363"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramStoryTop Influencer") Then
        Range("P" & Row).value = "721"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramStoryCelebrity") Then
        Range("P" & Row).value = "1842"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "InstagramStoryTop Celebrity") Then
        Range("P" & Row).value = "2218"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookPostNano") Then
        Range("P" & Row).value = "31"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookVideoNano") Then
        Range("P" & Row).value = "31"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookPostMicro") Then
        Range("P" & Row).value = "54"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookPostMega") Then
        Range("P" & Row).value = "86"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookPower") Then
        Range("P" & Row).value = "243"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookSuper Power") Then
        Range("P" & Row).value = "420"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookTop Influencer") Then
        Range("P" & Row).value = "650"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookCelebrity") Then
        Range("P" & Row).value = "792"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "FacebookTop Celebrity") Then
        Range("P" & Row).value = "1211"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "TwitterPostNano") Then
        Range("P" & Row).value = "31"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "TwitterPostMicro") Then
        Range("P" & Row).value = "54"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "TwitterPostMega") Then
        Range("P" & Row).value = "86"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
	ElseIf InStr(1, ActiveCell.value, "TwitterVideoMega") Then
        Range("P" & Row).value = "86"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
	ElseIf InStr(1, ActiveCell.value, "TwitterVideoMicro") Then
        Range("P" & Row).value = "54"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
	ElseIf InStr(1, ActiveCell.value, "TwitterVideoNano") Then
        Range("P" & Row).value = "31"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "TwitterPostPower") Then
        Range("P" & Row).value = "243"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "TwitterSuper Power") Then
        Range("P" & Row).value = "420"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "TwitterTop Influencer") Then
        Range("P" & Row).value = "650"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "TwitterCelebrity") Then
        Range("P" & Row).value = "792"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "TwitterTop Celebrity") Then
        Range("P" & Row).value = "1211"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeVideoNano") Then
        Range("P" & Row).value = "315"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeVideoMicro") Then
        Range("P" & Row).value = "555"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeVideoMega") Then
        Range("P" & Row).value = "850"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeVideoPower") Then
        Range("P" & Row).value = "1470"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeVideoSuper Power") Then
        Range("P" & Row).value = "2400"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeTop Influencer") Then
        Range("P" & Row).value = "3857"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeTop Influencer") Then
        Range("P" & Row).value = "3857"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeCelebrity") Then
        Range("P" & Row).value = "4122"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    ElseIf InStr(1, ActiveCell.value, "YouTubeTop Celebrity") Then
        Range("P" & Row).value = "6306"
        ActiveCell.NumberFormat = "General"
        Row = Row + 1
        Range(Col & Row).Activate
    Else
        Row = Row + 1
        Range(Col & Row).Activate
    End If
Wend
End Sub