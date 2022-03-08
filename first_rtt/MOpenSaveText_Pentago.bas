Attribute VB_Name = "MOpenSaveText_Pentago"


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


Public Sub OpenTextFile(fileName As String, Text1 As Object)  ' "DestTextBox As Object
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
Public Sub SaveTextFile(cmd_UL, cmd_UR, cmd_DL, cmd_DR As Object, fileName As String, which_player As Variant) ', mone_which_player As String)
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

' saving veribals

sContents = sContents + mone_which_player
 sContents = sContents + Chr$(13) & Chr$(10)
 
sContents = sContents + which_player
sContents = sContents + Chr$(13) & Chr$(10)


    ' כתיבת תוכנה של תיבת הטקס לקובץ
    Print #1, sContents
    Close #1  'L סגירת הקובץ
    
    ' במקרה של שגיאה
    If Err Then
        MsgBox Error, 48, "Dr VB"
    End If
    
End Sub

