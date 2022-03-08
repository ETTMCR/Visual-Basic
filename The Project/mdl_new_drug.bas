Attribute VB_Name = "mdl_new_drug"
Public go_drug_name  As String 'משתנה ששומר את שם התרופה שבחרנו לראות את פרטיה
Public go_comp_name  As String 'משתנה ששומר את שם החברה שבחרנו לראות את פרטיה וגם לצורך תרופה חדשה מאותה חברה
Public taking_comp As Boolean 'אם לחצת חברה קיימת הוספת תרופה חדשה יאפשר לקיחת שם החברה
Public taking_name_comp_2_drug As Boolean 'משתנה שמאפשר העברת שם חברה ממאגר חברות לתרופה
Public taking_name_comp_2_find_drug As Boolean 'משתנה שמאפשר העברת שם חברה ממאגר חברות לחיפוש תרופה

Public new_drug As Boolean
Public new_comp As Boolean
Public Sub which_to_use(which_use As String, ByVal use As ComboBox)
'מוסיף את סוגי ההתויות בהתאם לסוג השימוש הכללי"which"
use.AddItem Clear
Select Case which_use

Case "נגד דלקות":
use.Clear
use.AddItem "אחר"
use.AddItem "דלקות"
use.AddItem "דלקת גרון"
use.AddItem "דלקת עיניים"
use.AddItem "דלקת פרקים"
use.AddItem "הצטננות"

Case "תוספי מזון":
use.Clear
use.AddItem "אחר"
use.AddItem "תוספי ויטמינים"
use.AddItem "ויטמין A"
use.AddItem "ויטמין B"
use.AddItem "ויטמין C"
use.AddItem "ברזל"
use.AddItem "סידן"

Case "נגד כאבים":
use.Clear
use.AddItem "אחר"
use.AddItem "כאבים"
use.AddItem "כאב בטן"
use.AddItem "כאב גרון "
use.AddItem "כאב מחזור"
use.AddItem "כאב פרקים"
use.AddItem "כאב ראש "
use.AddItem "כאב שרירים"

Case "שונות":
use.Clear
use.AddItem "אחר"
use.AddItem "שלשול"
use.AddItem "חיטוי"
use.AddItem "הורדת לחץ דם"
use.AddItem "דילול דם"
use.AddItem "הורדת שומנים בדם"
use.AddItem "הורדת לחץ תוך עיני"
use.AddItem "נגד תולעי מעיים"
use.AddItem "אינסולין"
use.AddItem "פריחות"
use.AddItem "אלרגיות"

'Case "":
'Case "":
'use.AddItem ""

End Select

End Sub

Public Sub add_item_cbo_put_which_use(cbo As ComboBox)
'ממלא את הקומבו של שימושים כללי בערכים בעת פתיחת הטופס
cbo.Clear
 cbo.AddItem "נגד דלקות":
 cbo.AddItem "תוספי מזון":
cbo.AddItem "נגד כאבים"
cbo.AddItem "שונות":

End Sub

Public Sub ok_check_new_drug(name, comp, cost, which, lst_use, ByRef ok_n_d As Boolean, opt_type_else, Option1, Option2, Option3, Option4, Option5, fra_type)
'???? האם אפשר לשפוך את כל המשתנים מן הפריים מיד בלי לכתוב כל אחד
'בודק לקראת סיום האם כל השדות מלאים ותקינים ומודיע בהתאם
'במידה וכל הערכים תקינים
Dim opt As Boolean

If Not (name <> "") Then
MsgBox "לא כתבת שם תרופה", , " ! שים לב"
name.SetFocus
name.BackColor = &HFFC0FF
name = False
Else
name = True
End If

If Not (comp <> "") Then
MsgBox "לא כתבת שם חברה", , " ! שים לב"
comp.SetFocus
comp.BackColor = &HFFC0FF
comp = False
Else
comp = True
End If

If Not (cost <> "") Then
MsgBox "לא כתבת מחיר", , " ! שים לב"
cost.SetFocus
cost.BackColor = &HFFC0FF
cost = False
Else
cost = True
End If

If Not (which <> "") Then
MsgBox "לא ציינת סוג שימוש כללי", , " ! שים לב"
which.SetFocus
which.BackColor = &HFFC0FF
which = False
Else
which = True
End If


If (lst_use.ListCount = 0) Then
MsgBox "לא ציינת סוג שימוש ", , " ! שים לב"
lst_use.SetFocus
lst_use.BackColor = &HFFC0FF
lst_use = False
Else
lst_use = True
End If

If (opt_type_else.Value = 0) And (Option1.Value = 0) And (Option2.Value = 0) _
    And (Option3.Value = 0) And (Option4.Value = 0) And (Option5.Value = 0) Then
        MsgBox "לא ציינת סוג תכשיר ", , " ! שים לב"
        fra_type.BackColor = &HFFC0FF
        opt_ok = False
    Else
        opt_ok = True
End If

If name And comp And cost And lst_use And which And opt_ok Then
ok_n_d = True
Else
ok_n_d = False
End If
End Sub
