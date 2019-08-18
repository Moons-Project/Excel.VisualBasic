' 这是一个单文件的VisualBasic演练场
' 本质上是需要在Excel中使用VisualBasic，所以至少需要先了解一些基础的语法啥的
' reference https://docs.microsoft.com/zh-cn/dotnet/visual-basic/programming-guide/language-features/

' 关于如何运行
' 可以打开PowerShell.exe, 并键入
'   vbc .\.VisualBasic.vb ; .\.VisualBasic.exe



' 一些通用的方法
Module Tools
  ' Debug Flag
  Function DebugFlag() As Boolean
    Return true
  End Function

  ' Log Function
  Sub Print(value As String)
    if (DebugFlag())
      System.Console.WriteLine(value)
    End if
  End Sub

  ' Assert
  Sub Assert(Of T As IComparable)(ByVal T1 As T, ByVal T2 As T)
    ' System.Diagnostics.Debug.Assert(value, "Error")
    if (Not T1.CompareTo(T2) = 0)
      Throw New System.Exception($"{T1} != {T2}, Assertion Failed")
    End if
  End Sub


  ' 退出逻辑
  Function ExitPlayGround() As Boolean
    ' Console.Write("Press Any key to continue..")
    ' Console.ReadKey(true)
    Return true
  End Function

End Module

' 主程序
Module Program

  Sub PartArray()
    Tools.Print("Array Part")

    ' 花式声明
    Dim numbersExample_1(4) As Integer ' 说是4，其实有5个元素
    Dim numbersExample_2 = new String() {"0", "1", "2", "3", "4"}
    Redim numbersExample_1(5) ' 改变大小，不保留值
    ReDim Preserve numbersExample_2(3) ' 改变大小，保留值
    Dim matrixExample_1(4, 4) As Double ' 6x6 2维数组
    Dim matrixExample_2 = new Integer(1, 1) {{1, 2}, {3, 4}}
    Dim matrixExample_3()() As Integer = New Integer(2)() {} ' jagged / 钩子数组
    ' 大概是这么着用
    matrixExample_3(0) = numbersExample_1

    ' 访问
    For index As Integer = 0 to numbersExample_2.GetUpperBound(0) ' 维度0最高缩引
      ' Tools.Print($"index: {index}, Value: {numbersExample_1(index)}")
      Tools.Assert(Of String)($"{index}", numbersExample_2(index))
    Next

    ' 大小
    Tools.Assert(Of Integer)(matrixExample_1.Length, 
      (matrixExample_1.GetUpperBound(0) + 1)* 
      (matrixExample_1.GetUpperBound(1) + 1))


    ' 类型
    Dim typeArrayExample As Array = Array.CreateInstance(GetType(Object), 1)
    Tools.Assert(Of String)(typeArrayExample.GetType().Name, "Object[]")


  end Sub


  Sub PartProcess()
    ' 过程
    ' 1. Function 和 Sub 的区别是前者有返回值，后者没有

  End Sub


  ' Async Function AsyncExample(num as Integer) As _
  '   System.Threading.Tasks.Task(Of Integer)
  '   ' Await System.Threading.Tasks.Task.Run()
  '   Tools.Print(num)
  '   Return num
  ' End Function


  ' 程序入口
  Sub Main(args As String())
    PartArray()
    PartProcess()

    ' AsyncExample(1)

    Tools.ExitPlayGround()
  End Sub
End Module