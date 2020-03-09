VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'mmap
'Copyright (c) 2020 mmYYmmdd
Option Explicit

    Private Declare PtrSafe Function write_mmap Lib "msofficeutil.dll" (ByRef data_ As Variant, ByRef name_ As Variant, ByRef bytes_ As Long) As LongPtr
    Private Declare PtrSafe Function reserve_mmap Lib "msofficeutil.dll" (ByVal bytes_ As Long, ByRef name_ As Variant) As LongPtr
    Private Declare PtrSafe Sub closeHandle Lib "msofficeutil.dll" (ByVal h As LongPtr)
    Private Declare PtrSafe Sub read_mmap Lib "msofficeutil.dll" (ByRef name_ As Variant, ByRef out As Variant, ByVal bytes_ As Long)
    
    Private nameFor As String
    Private handleFor As LongPtr
    Private sizeFor As Long
'******************************************************************************************
' JSONフォーマットでMMF（メモリマップドファイル）に格納したVariant変数をラップしたクラス
' ユーザーはこのクラスを直接インスタンス化せず
' mmapFactoryモジュールの writeMMF/reserveMMF を経由してインスタンスを作成する
'******************************************************************************************

' 自分の管理するMMF内のJSONの内容をVariant変数に書き込む
' 事前に out = loads() と Set out = loads() のどちらにすべきか不明なためSub
Public Sub loads(ByRef out As Variant)
    Call read_mmap(nameFor, out, 0)
End Sub

Public Property Get Name() As String
    Name = nameFor
End Property

Public Property Get Size() As Long
    Size = sizeFor
End Property


'*********************************************************
' ユーザーは以下の Function/Sub を使用しない

    Private Sub Class_Initialize()
        nameFor = ""
        handleFor = 0
        sizeFor = 0
    End Sub

    Private Sub Class_Terminate()
        If handleFor <> 0 Then
            Call closeHandle(handleFor)
        End If
    End Sub
'*********************************************************

    ' 一度しかできない
    Function dumps(ByRef data_) As Boolean
        If handleFor = 0 Then
            nameFor = "A_" & for_Class_Initialize()
            handleFor = write_mmap(data_, nameFor, sizeFor)
            dumps = (handleFor <> 0)
        Else
            dumps = False
        End If
    End Function
    
    ' 一度しかできない
    Function reserve_buffer(ByVal bytesize As Long) As Boolean
        If handleFor = 0 And 0 < bytesize Then
            nameFor = "R_" & for_Class_Initialize()
            handleFor = reserve_mmap(bytesize, nameFor)
            If handleFor <> 0 Then sizeFor = bytesize
            reserve_buffer = (handleFor <> 0)
        Else
            reserve_buffer = False
        End If
    End Function
