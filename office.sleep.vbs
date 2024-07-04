Sub sleep(T As Single)  ' T 参数的单位是 秒级
    Dim time1 As Single
    time1 = Timer
    Do
        DoEvents '转让控制权，以便让操作系统处理其它的事件
    Loop While Timer - time1 < T  ' T 参数的单位是 秒级
End Sub