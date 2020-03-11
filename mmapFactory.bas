Attribute VB_Name = "mmapfactory"
'mmapFactory
'Copyright (c) 2020 mmYYmmdd
Option Explicit
    
    Private Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
    End Type
    
' Variant --> JSON
Public Declare PtrSafe Function VariantToJson Lib "msofficeutil.dll" (ByRef data_ As Variant) As Variant
' JSON --> Variant（事前に out = loads() と Set out = loads() のどちらにすべきか不明なためSub）
Public Declare PtrSafe Sub JsonToVariant Lib "msofficeutil.dll" (ByRef expr As Variant, ByRef out As Variant)

    Declare PtrSafe Function SetCurrentDirectoryA Lib "Kernel32" (ByVal pathName As String) As Long
    Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.dll" (ByRef g As GUID) As Long

' pythonのmain関数の引数としてのMMFオブジェクト生成
Function dumpToMMF(ByRef data_ As Variant) As mmap
    Set dumpToMMF = New mmap
    If Not dumpToMMF.dumps(data_) Then Set dumpToMMF = Nothing
End Function
    
' pythonのmain関数の引数（返り値格納用）としてのMMFオブジェクト生成
' （JSON表現として十分な大きさをバイト数で与える）
Function reserveMMF(ByVal bytesize As Long) As mmap
    Set reserveMMF = New mmap
    If Not reserveMMF.reserve_buffer(bytesize) Then reserveMMF = Nothing
End Function

' pythonスクリプトの実行（返り値の取得）
' ParameterS のうちひとつだけ reserveMMF で作成したオブジェクト
'（スクリプトファイルの存在するフォルダからPythonが起動できる必要がある「アプリ実行エイリアス」に注意！）
Function execPythonModule(ByVal scriptPath As String, ParamArray ParameterS() As Variant) As Variant
    Dim arguments As String, z As Variant
    Dim ReturnBuffer As mmap:   Set ReturnBuffer = Nothing
    For Each z In ParameterS
        If TypeName(z) = "mmap" Then
            arguments = arguments & " " & z.Name & "+" & z.Size
            If Left(z.Name, 1) = "R" Then Set ReturnBuffer = z
        Else
            arguments = arguments & " " & z
        End If
    Next z
    '--------------------------------------------------------
    Dim wSS As Object, fileAttr As Variant
    Set wSS = CreateObject("WScript.Shell")
    fileAttr = getFileAttr(scriptPath)      '(folder, name)
    Dim curDir_ As String:  curDir_ = CurDir()
    SetCurrentDirectoryA fileAttr(0)
    Dim command As String
    command = "python " & fileAttr(1) & arguments
    Call wSS.Run(command, 0, True)
    SetCurrentDirectoryA curDir_
    If Not ReturnBuffer Is Nothing Then
        Call ReturnBuffer.loads(execPythonModule)   '返り値の取得
    End If
End Function


    'for mmap Class（mmapクラスインスタンスの初期化時のオブジェクト名）
    Public Function for_Class_Initialize() As String
        Static guid_ As String
        If Len(guid_) = 0 Then guid_ = getGUID_()
        Static counter_ As Long
        counter_ = counter_ + 1
        for_Class_Initialize = guid_ & "_" & counter_
    End Function

    '(FolderName, FileName)
    Private Function getFileAttr(ByVal path As String) As Variant
        Dim fso As Object, fileObject As Object
        Set fso = CreateObject("Scripting.FIleSystemObject")
        Set fileObject = fso.GetFile(path)
        getFileAttr = VBA.Array(fileObject.ParentFolder.path, fileObject.Name)
    End Function
    
    Private Function getGUID_() As String
        Dim g As GUID
        Call CoCreateGuid(g)
        getGUID_ = Hex(g.Data1) & Hex(g.Data2) & Hex(g.Data3) _
                & Hex(g.Data4(0)) & Hex(g.Data4(1)) & Hex(g.Data4(2)) & Hex(g.Data4(3)) _
                & Hex(g.Data4(4)) & Hex(g.Data4(5)) & Hex(g.Data4(6)) & Hex(g.Data4(7))
    End Function
    
    
