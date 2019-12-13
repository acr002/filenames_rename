'-----------------------------------------------------------[date: 2019.12.09]
Attribute VB_Name = "rename"
Option Explicit

'***********************************************
' rename
' 2019.11.27(水).new
'***********************************************
' 別の拡張子に変更することはできません。なんとかした方がいいかも。
'***********************************************
Private fso As Object
'***********************************************
Private Enum header_pj
  hpj_seq = 1
  hpj_old_dir = 2
  hpj_old_file_name = 3
  hpj_flag_on_move = 4
  hpj_new_dir = 5
  hpj_new_file_name = 6
  hpj_remarks = 7
End Enum
'***********************************************
Private Type elements
  dir   As String
  path  As String
  fname As String
  fex   As String
End Type
'***********************************************

Public Sub main_rename()
  Dim pass_line  As Boolean
  Dim buf        As String
  Dim start_time As Single
  Dim a          As elements
  Dim b          As elements
  Dim ws         As Worksheet
  Dim y          As Long
  Dim ys         As Long
  Dim pj         As Variant
  start_time = Timer
  Set fso = Createobject("Scripting.FileSystemObject")
  Set ws = ThisWorkbook.worksheets(1)
  pj = ws.usedrange.value
  ys = UBound(pj, 1)
  For y = 2 To ys
    If Len(pj(y, hpj_old_dir)) = 0 Then
      pass_line = True
    End If
    If Len(pj(y, hpj_old_file_name)) = 0 Then
      pass_line = True
    End If
    If pass_line Then
      pass_line = False
    Else
      a.dir = setup_dir(trim(pj(y, hpj_old_dir)))
      If Len(trim(pj(y, hpj_new_dir))) Then
        b.dir = setup_dir(trim(pj(y, hpj_new_dir)))
      Else
        b.dir = a.dir
      End If
      a.fname = pj(y, hpj_old_file_name)
      b.fname = pj(y, hpj_new_file_name)
      If Len(trim(b.fname)) = 0 Then
        b.fname = a.fname
      End If
      a.fex = split_ex(a.fname)
      b.fex = split_ex(b.fname)
      If a.fex <> b.fex Then
        b.fex = a.fex
        b.fname = b.fname & "." & b.fex
      End If
      If fso.folderexists(a.dir) Then
        a.path = a.dir & a.fname
        If fso.fileexists(a.path) Then
          b.path = b.dir & b.fname
          If Not fso.fileexists(b.path) Then
            Call check_path(b.dir)
            If pj(y, hpj_flag_on_move) = 1 Then
              fso.movefile Source:=a.path, Destination:=b.path
              buf = "move"
            Else
              fso.copyfile Source:=a.path, Destination:=b.path
              buf = "copy"
            End If
          Else
            buf = "未処理(移動先に同名のファイルあります)"
          End If
        Else
          buf = "未処理(指定のファイルがありません)"
        End If
      Else
        buf = "未処理(指定のフォルダがありません)"
      End If
    End If
    If Len(buf) Then
      ws.cells(y, hpj_remarks).value = buf
      buf = ""
    Else
      ws.cells(y, hpj_remarks).value = "指定に抜けがあります"
    End If
  Next y
  MsgBox "end of run" & vbCrLf & "time:" & CStr(Timer - start_time)
End Sub
'-----------------------------------------------------------------------------

Private Sub check_path(ByVal out_path As String)
  Dim ar As Variant
  Dim t_path As String
  Dim i As Long
  If Not fso.folderexists(out_path) Then
    ar = split(out_path, "\")
    t_path = ar(0) & "\"
    For i = 1 To UBound(ar) - 1
      t_path = t_path & ar(i) & "\"
      If Not fso.folderexists(t_path) Then
        fso.createfolder t_path
      End If
    Next i
  End If
End Sub
'-----------------------------------------------------------------------------

Private Function setup_dir(ByVal a_path As String) As String
  If Right(a_path, 1) <> "\" Then
    setup_dir = a_path & "\"
  Else
    setup_dir = a_path
  End If
End Function
'-----------------------------------------------------------------------------

Private Function split_ex(filename As String) As String
  Dim ar As Variant
  ar = split(filename, ".")
  split_ex = ar(UBound(ar))
End Function
'-----------------------------------------------------------------------------

