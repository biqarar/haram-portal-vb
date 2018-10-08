VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function TarikhShamsi(Optional date1 As String, Optional SmallDate1 As Boolean) As String

      '====================================================
      Dim d, p, w, mon, Mm, Ym, u, v, rp, x, i, Ys, Ms, Dm, P1, D1, Ds, DateShamsi
      d = Array(20, 19, 20, 20, 21, 21, 22, 22, 22, 22, 21, 21)
      p = Array(11, 12, 10, 12, 11, 11, 10, 10, 10, 9, 10, 10)
      w = Array("Ìò‘‰»Â", "œÊ‘‰»Â", "”Â ‘‰»Â", "çÂ«—‘‰»Â", "Å‰Ã‘‰»Â", "Ã„⁄Â", "‘‰»Â")
      
      If SmallDate1 = True Then
            mon = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
      Else
            mon = Array("›—Ê—œÌ‰", "«—œÌ»Â‘ ", "Œ—œ«œ", " Ì—", "„—œ«œ", "‘Â—ÌÊ—", "„Â—", "¬»«‰", "¬–—", "œÌ", "»Â„‰", "«”›‰œ")
      End If
      
      If date1 = "" Then date1 = Date
      
      Dm = Day(date1) '»œ”  ¬Ê—œ‰ —Ê“
      Mm = Month(date1) '»œ”  ¬Ê—œ‰ „«Â
      Ym = Year(date1) '»œ”  ¬Ê—œ‰ ”«·
      u = 0
      rp = 0
      If (Ym Mod 4) = 0 Then u = 1 ' ‘ŒÌ’ ò»Ì”Â »Êœ‰
      If ((Ym Mod 100) = 0 And (Ym Mod 400) <> 0) Then u = 0 ' ‘ŒÌ’ ò»Ì”Â ‰»Êœ‰
      Ys = Ym - 622 ' »œÌ· ”«· „Ì·«œÌ »Â ‘„”Ì
      x = Ys - 22
      x = x Mod 33
      If ((x Mod 4) = 0 And x <> 32) Then rp = 1
      i = Not (rp - 2) + Not (u - 2) * 2
      x = 0
      If (i = 0 And Mm = 3) Then x = 1
      If i = 0 Then i = 3
      Ms = (9 + Mm) Mod 13
      If Ms < 10 Then Ms = Ms + 1
      D1 = d(Mm - 1)
      If (i = 1 And Mm > 2) Then D1 = D1 - 1
      If (i = 2 And Mm < 3) Then D1 = D1 - 1
      P1 = p(Mm - 1)
      If (i = 1 And Mm > 2) Then P1 = P1 + 1
      If (i = 2 And Mm < 4) Then P1 = P1 + 1
      If (Dm > 0 And Dm <= D1) Then
             Ds = P1 + Dm + x - 1
          x = 1
      Else
          Ds = Dm - D1
          Ms = Ms + 1
          If Ms = 13 Then Ms = 1
          x = 2
      End If
      If ((Mm = 3 And x = 2) Or Mm > 3) Then Ys = Ys + 1
      If SmallDate1 = True Then
'     ??? ??? ?? ???? ???? ???????? ???????? ?? ??? ?? ?? ???? ????? ?? ?????
'            TarikhShamsi = Trim(Str(Ys)) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(Str(Ds))
            TarikhShamsi = Mid(Trim(Str(Ys)), 3, 2) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(Str(Ds))
      Else
            TarikhShamsi = w(Weekday(Date) - 1) + " " + Str(Ds) + " " + mon(Ms - 1) + " " + Str(Ys)
      End If

End Function

Private Sub Form_Load()
MsgBox TarikhShamsi(Date)
Form1.Caption = TarikhShamsi(Date)
End Sub
