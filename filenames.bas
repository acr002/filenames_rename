'-----------------------------------------------------------[date: 2019.12.02]
Attribute VB_Name = "Module1"
Option Explicit

'***********************************************
' 2019.11.27(水).new
' 2019.11.28(木)  修正(inフォルダへのダイレクトをやめました。
'***********************************************
private fso as Object
'***********************************************

Public Sub main_get_filenames()
  Dim t_folder  As string
  Dim col       As Collection
  Dim dir_cur   As String
  Dim drive_cur As String
  Dim ar        As Variant
  Dim i         As Long
  Dim buf       As String
  Dim y         As Long
  Dim ws        As Worksheet
  Dim b_fn      As Variant
  Set fso = Createobject("Scripting.FileSystemObject")
  dir_cur = ThisWorkbook.Path & "\"
  't_folder = dir_cur & "in\"
  'if not fso.folderexists(t_folder) then
    ar = Split(dir_cur, "\")
    drive_cur = ar(LBound(ar))
    ChDrive drive_cur
    ChDir dir_cur
    With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then
        t_folder = .SelectedItems(1)
      else
        t_folder = dir_cur
      End If
    End With
  'end if
  set col = get_filenames_sub(t_folder)
  y = 1
  Set ws = ThisWorkbook.Worksheets(1)
  For Each b_fn In col
    y = y + 1
    ar = Split(b_fn, "\")
    For i = LBound(ar) To UBound(ar) - 1
      buf = buf & ar(i) & "\"
    Next i
    ws.Cells(y, 1).Value = y - 1
    ws.Cells(y, 2).Value = buf
    ws.Cells(y, 3).Value = ar(UBound(ar))
    buf = ""
  Next b_fn
  MsgBox "end of run"
End Sub
'-----------------------------------------------------------------------------

Private Function get_filenames_sub(ByVal a_path As String) As Collection
  Dim r_cc     As Collection
  Dim cc       As Collection
  Dim ii       As Variant
  Dim b_file   As Object
  Dim b_folder As Object
  Set cc = New Collection
  For Each b_file In fso.getfolder(a_path).files
    cc.Add b_file.Path
  Next b_file
  For Each b_folder In fso.getfolder(a_path).subfolders
    Set r_cc = get_filenames_sub(b_folder.Path)
    For Each ii In r_cc
      cc.Add ii
    Next ii
  Next b_folder
  Set get_filenames_sub = cc
End Function
'-----------------------------------------------------------------------------

