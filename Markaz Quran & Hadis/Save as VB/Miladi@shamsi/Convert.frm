VERSION 5.00
Begin VB.Form Taqvim 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6270
   Icon            =   "Convert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.Label KKK 
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Tarikh 
      Height          =   975
      Left            =   1440
      TabIndex        =   1
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1575
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "Taqvim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function TarikhShamsi(Optional date1 As String, Optional SmallDate1 As Boolean) As String

      '====================================================
      Dim d, p, w, mon, mm, Ym, u, v, rp, X, I, Ys, Ms, Dm, P1, D1, Ds, DateShamsi
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
      mm = Month(date1) '»œ”  ¬Ê—œ‰ „«Â
      Ym = Year(date1) '»œ”  ¬Ê—œ‰ ”«·
      u = 0
      rp = 0
      If (Ym Mod 4) = 0 Then u = 1 ' ‘ŒÌ’ ò»Ì”Â »Êœ‰
      If ((Ym Mod 100) = 0 And (Ym Mod 400) <> 0) Then u = 0 ' ‘ŒÌ’ ò»Ì”Â ‰»Êœ‰
      Ys = Ym - 622 ' »œÌ· ”«· „Ì·«œÌ »Â ‘„”Ì
      X = Ys - 22
      X = X Mod 33
      If ((X Mod 4) = 0 And X <> 32) Then rp = 1
      I = Not (rp - 2) + Not (u - 2) * 2
      X = 0
      If (I = 0 And mm = 3) Then X = 1
      If I = 0 Then I = 3
      Ms = (9 + mm) Mod 13
      If Ms < 10 Then Ms = Ms + 1
      D1 = d(mm - 1)
      If (I = 1 And mm > 2) Then D1 = D1 - 1
      If (I = 2 And mm < 3) Then D1 = D1 - 1
      P1 = p(mm - 1)
      If (I = 1 And mm > 2) Then P1 = P1 + 1
      If (I = 2 And mm < 4) Then P1 = P1 + 1
      If (Dm > 0 And Dm <= D1) Then
             Ds = P1 + Dm + X - 1
          X = 1
      Else
          Ds = Dm - D1
          Ms = Ms + 1
          If Ms = 13 Then Ms = 1
          X = 2
      End If
      If ((mm = 3 And X = 2) Or mm > 3) Then Ys = Ys + 1
      If SmallDate1 = True Then
'     ??? ??? ?? ???? ???? ???????? ???????? ?? ??? ?? ?? ???? ????? ?? ?????
'            TarikhShamsi = Trim(Str(Ys)) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(Str(Ds))
            TarikhShamsi = Mid(Trim(Str(Ys)), 3, 2) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(Str(Ds))
           If Val(Ys) < 10 Then Ys = "0" & Val(Ys)
           If Val(Ms) < 10 Then Ms = "0" & Val(Ms)
           If Val(Ds) < 10 Then Ds = "0" & Val(Ds)
            
            Tarikh.Caption = Val(Ys) & "/" & (Ms) & "/" & Val(Ds)
            ' Tarikh.Caption = Ys & Ms & Ds
      Else
            TarikhShamsi = w(Weekday(Date) - 1) + " " + Str(Ds) + " " + mon(Ms - 1) + " " + Str(Ys)
           If Val(Ys) < 10 Then Ys = "0" & Val(Ys)
           If Val(Ms) < 10 Then Ms = "0" & Val(Ms)
           If Val(Ds) < 10 Then Ds = "0" & Val(Ds)
            
            Tarikh.Caption = Val(Ys) & "/" & (Ms) & "/" & Val(Ds)
            'Tarikh.Caption = Ys & Ms & Ds
            KKK.Caption = Ys & Ms & Ds
      End If

End Function

Private Sub Form_Load()
'MsgBox TarikhShamsi(Date)
'Form1.Caption = TarikhShamsi(Date)
Label1.Caption = TarikhShamsi(Date)
End Sub

