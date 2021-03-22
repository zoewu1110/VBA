Attribute VB_Name = "Module1"
Sub ChatRobot()

    MsgBox "歡迎來到雲寶寶壽司店"
    
    Dim userName As String
    userName = InputBox("請問你叫甚麼名字?")
    MsgBox "Hi!" & userName, 0, "歡迎歡迎~!!!!!"
    
    Dim userAge As Integer
    userAge = InputBox("偷偷給訴我你幾歲好嗎?")
    If userAge > 20 Then
        MsgBox "真的假的啊~姐姐好!"
    Else
        MsgBox "我想我們可以做很好的朋友"
    End If

    Dim rst As Integer
    rst = MsgBox("雲寶寶今年三歲，你喜歡吃壽司嗎?", 4, "壽司狂熱已上線")
    If vbYes Then
        MsgBox "好開勳喔~", 0, "壽司狂熱已上線"
    Else
        MsgBox "哭哭QQ", 0, "壽司狂熱已下線"
    End If
    
    
    
End Sub
