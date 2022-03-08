Attribute VB_Name = "MOpenSaveText"
'Option Explicit

Public Sub Unchecked_all(cmd_chk1, cmd_chk2, cmd_chk3, cmd_chk4, cmd_chk5, cmd_chk6, cmd_chk7, cmd_chk8, cmd_chk9, cmd_chk10, cmd_chk11, cmd_chk12, cmd_chk13, cmd_chk14, cmd_chk15, cmd_chk16, cmd_chk17, cmd_chk18, cmd_chk19, cmd_chk20, cmd_chk21, cmd_chk22, cmd_chk23, cmd_chk24, cmd_chk25, Label26 As Object)
Dim h As Integer
'מאפס את התאים

For h = 0 To 24

cmd_chk1(h).BackColor = &HFFFFFF
cmd_chk2(h).BackColor = &HFFFFFF
cmd_chk3(h).BackColor = &HFFFFFF
cmd_chk4(h).BackColor = &HFFFFFF
cmd_chk5(h).BackColor = &HFFFFFF
cmd_chk6(h).BackColor = &HFFFFFF
cmd_chk7(h).BackColor = &HFFFFFF
cmd_chk8(h).BackColor = &HFFFFFF
cmd_chk9(h).BackColor = &HFFFFFF
cmd_chk10(h).BackColor = &HFFFFFF
cmd_chk11(h).BackColor = &HFFFFFF
cmd_chk12(h).BackColor = &HFFFFFF
cmd_chk13(h).BackColor = &HFFFFFF
cmd_chk14(h).BackColor = &HFFFFFF
cmd_chk15(h).BackColor = &HFFFFFF
cmd_chk16(h).BackColor = &HFFFFFF
cmd_chk17(h).BackColor = &HFFFFFF
cmd_chk18(h).BackColor = &HFFFFFF
cmd_chk19(h).BackColor = &HFFFFFF
cmd_chk20(h).BackColor = &HFFFFFF
cmd_chk21(h).BackColor = &HFFFFFF
cmd_chk22(h).BackColor = &HFFFFFF
cmd_chk23(h).BackColor = &HFFFFFF
cmd_chk24(h).BackColor = &HFFFFFF
cmd_chk25(h).BackColor = &HFFFFFF

Next

For h = 0 To 324 ' null all label26
Label26(h).Caption = ""
Next
'מאפס את התאים סיום
End Sub


'@@@@@@@@@@@@@@@@@

'Dim sString As String, sInput As String
'Open "C:\File.txt" For Input As #txtFile
'While Not EOF(txtFile)
 '  Input #txtFile, sInput
  ' sString = sString & sInput
   'sInput = ""
'Wend
'Close #txtFile
'@@@@@@@@@@@@@@@@@


Public Sub OpenTextFile(DestTextBox As Object, fileName As String, Text1 As Object)
Dim i As Integer
Dim sString As String, sInput As String

    Dim fIndex As Integer
    
    On Error Resume Next
    ' פתיחת הקובץ אותו בחר המשתמש
    Open fileName For Input As #1
    ' במקרה של שגיאה
    If Err Then
        MsgBox "Can't open file: " + fileName, vbOKOnly, "Dr VB"
        Exit Sub
    End If
    ' הצגת הסמן בצורת שעון-החול בזמן ביצוע השמירה
    Screen.MousePointer = 11
    ' קבלת תוכן הקובץ לתיבת הטקסט
    '! DestTextBox.Text = Input(LOF(1), 1)
    
    'for i                                          'המספר אם רצים בתוים תמיד יהיה 50 כי יש 0 או1 תמיד
    
    
    '@@@@@@@@@@@@@@@
    While Not EOF(1)
   Input #1, sInput
   sString = sString & sInput '+ "5" 'Chr$(13) & Chr$(10)
   sInput = ""

Wend
Text1.Text = sString
    '@@@@@@@@@@@@@@@
    
    Close #1
    ' החזרת הסמן למצבו ברגיל
    Screen.MousePointer = 0
    
End Sub
Public Sub SaveTextFile(cmd_UL, cmd_UR, cmd_DL, cmd_DR As Object, fileName As String)
Dim i As Integer

' פרוצדורה לשמירת תוכנה של תיבת טקסט
' הארגומנט הראשון הוא תיבת הטקסט שאת תוכנה רוצים לשמור
' הארגומנט השני הוא שם הקובץ שאליו לשמור את תוכנה של תיבת הטקסט

    On Error Resume Next
    Dim sContents As String

    ' פתיחת קובץ לשמירה
    Open fileName For Output As #1
    
   'UL
For i = 0 To 8 ' could one long string
      'sContents = sContents + CStr(cmd_UL(i).BackColor) + " "
      If cmd_UL(i).BackColor = -2147483630 Then
          sContents = sContents + CStr(0)
       Else
            If cmd_UL(i).BackColor = 16777215 Then
                    sContents = sContents + CStr(1)
           Else
                     sContents = sContents + CStr(2)
          End If
      End If '
Next
sContents = sContents + Chr$(13) & Chr$(10)
         
   'UR
For i = 0 To 8 ' could one long string
      'sContents = sContents + CStr(cmd_UL(i).BackColor) + " "
      If cmd_UR(i).BackColor = -2147483630 Then
          sContents = sContents + CStr(0)
       Else
            If cmd_UR(i).BackColor = 16777215 Then
                    sContents = sContents + CStr(1)
           Else
                     sContents = sContents + CStr(2)
          End If
      End If '
Next
sContents = sContents + Chr$(13) & Chr$(10)
         
   'DL
For i = 0 To 8 ' could one long string
      'sContents = sContents + CStr(cmd_UL(i).BackColor) + " "
      If cmd_DL(i).BackColor = -2147483630 Then
          sContents = sContents + CStr(0)
       Else
            If cmd_DL(i).BackColor = 16777215 Then
                    sContents = sContents + CStr(1)
           Else
                     sContents = sContents + CStr(2)
          End If
      End If '
Next
sContents = sContents + Chr$(13) & Chr$(10)
   
'DR
For i = 0 To 8 ' could one long string
      'sContents = sContents + CStr(cmd_UL(i).BackColor) + " "
      If cmd_DR(i).BackColor = -2147483630 Then
          sContents = sContents + CStr(0)
       Else
            If cmd_DR(i).BackColor = 16777215 Then
                    sContents = sContents + CStr(1)
           Else
                     sContents = sContents + CStr(2)
          End If
      End If '
Next
sContents = sContents + Chr$(13) & Chr$(10)


    ' כתיבת תוכנה של תיבת הטקס לקובץ
    Print #1, sContents
    Close #1  'L סגירת הקובץ
    
    ' במקרה של שגיאה
    If Err Then
        MsgBox Error, 48, "Dr VB"
    End If
    
End Sub

