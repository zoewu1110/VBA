Attribute VB_Name = "Module1"
Sub ChatRobot()

    MsgBox "�w��Ө춳�_�_�إq��"
    
    Dim userName As String
    userName = InputBox("�аݧA�s�ƻ�W�r?")
    MsgBox "Hi!" & userName, 0, "�w���w��~!!!!!"
    
    Dim userAge As Integer
    userAge = InputBox("�������D�ڧA�X���n��?")
    If userAge > 20 Then
        MsgBox "�u��������~�j�j�n!"
    Else
        MsgBox "�ڷQ�ڭ̥i�H���ܦn���B��"
    End If

    Dim rst As Integer
    rst = MsgBox("���_�_���~�T���A�A���w�Y�إq��?", 4, "�إq�g���w�W�u")
    If vbYes Then
        MsgBox "�n�}����~", 0, "�إq�g���w�W�u"
    Else
        MsgBox "����QQ", 0, "�إq�g���w�U�u"
    End If
    
    
    
End Sub
