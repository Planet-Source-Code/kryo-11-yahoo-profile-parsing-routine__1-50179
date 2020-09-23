VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Yahoo Profile Parsing"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   443
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Get Profile"
      Default         =   -1  'True
      Height          =   360
      Left            =   4680
      TabIndex        =   39
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   495
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   4800
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   495
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   495
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   3600
      Width           =   4935
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   6390
      Width           =   6735
      Begin VB.Label Status 
         AutoSize        =   -1  'True
         Caption         =   "Idle"
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4455
   End
   Begin VB.Line Line4 
      X1              =   6
      X2              =   6
      Y1              =   30
      Y2              =   8
   End
   Begin VB.Line Line3 
      X1              =   306
      X2              =   306
      Y1              =   30
      Y2              =   8
   End
   Begin VB.Line Line2 
      X1              =   6
      X2              =   306
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000012&
      X1              =   6
      X2              =   306
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   14
      Left            =   1680
      TabIndex        =   38
      Top             =   6120
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   13
      Left            =   1680
      TabIndex        =   37
      Top             =   5880
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   12
      Left            =   1680
      TabIndex        =   36
      Top             =   5640
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   1680
      TabIndex        =   35
      Top             =   5400
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cool Link 2:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   17
      Left            =   0
      TabIndex        =   34
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cool Link 1:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   16
      Left            =   0
      TabIndex        =   33
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Homepage:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   15
      Left            =   0
      TabIndex        =   32
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cool Link 3:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   14
      Left            =   0
      TabIndex        =   31
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Favorite Quote:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   13
      Left            =   0
      TabIndex        =   29
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Latest News:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   12
      Left            =   0
      TabIndex        =   27
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Hobbies:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   11
      Left            =   0
      TabIndex        =   25
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   1680
      TabIndex        =   24
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   1680
      TabIndex        =   23
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   1680
      TabIndex        =   22
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   21
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   20
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   19
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   18
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   17
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   16
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   15
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   14
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Occupation:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Yahoo! ID:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Real Name:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nickname:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Location:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Age:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Marital Status:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Gender:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Email:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Member Since:"
      ForeColor       =   &H80000015&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type YAHOO_PROFILE_INFORMATION
    MemberSince As String
    LastUpdated As String
    Email As String
    YahooID As String
    RealName As String
    NickName As String
    Location As String
    Age As String
    MaritalStatus As String
    Gender As String
    Occupation As String
    Hobbies As String
    LatestNews As String
    FavoriteQuote As String
    HomePage As String
    CoolLink1 As String
    CoolLink2 As String
    CoolLink3 As String
End Type

Dim NewProfile As YAHOO_PROFILE_INFORMATION

Private Sub cmdCheck_Click()
    If Inet1.StillExecuting Then
        Status = "Still executing last request.  Please wait."
        Exit Sub
    End If
    Dim RetData, tempData
    Status = "Connecting to users profile."
    RetData = Inet1.OpenURL("http://profiles.yahoo.com/" & txtName)
    Status = "Parsing profile information."
    If InStr(RetData, "Sorry, but the profile you are looking for is not currently available") Then GoTo NoProfile
    If RetData = "" Then GoTo NoProfile
    ParseYahooProfile RetData
    Status = "Idle"
    With NewProfile
        Label2(0) = .MemberSince:   Label2(1) = .LastUpdated
        Label2(2) = .Email:         Label2(3) = .YahooID
        Label2(4) = .RealName:      Label2(5) = .NickName
        Label2(6) = .Location:      Label2(7) = .Age
        Label2(8) = .MaritalStatus: Label2(9) = .Gender
        Label2(10) = .Occupation:   Label2(11) = .HomePage
        Label2(12) = .CoolLink1:    Label2(13) = .CoolLink2
        Label2(14) = .CoolLink3
        Text1 = .Hobbies: Text2 = .LatestNews: Text3 = .FavoriteQuote
    End With
    Exit Sub
NoProfile:
    Status = "Sorry, but the profile you are looking for is not currently available"
    Exit Sub
End Sub

Sub ParseYahooProfile(Data)
    Dim BackUP
    Dim LeftPos As Integer
    Dim RightPos As Integer
    With NewProfile
        .Age = "": .CoolLink1 = "": .CoolLink2 = "": .CoolLink3 = "": .Email = ""
        .FavoriteQuote = "": .Gender = "": .Hobbies = "": .HomePage = "": .LastUpdated = ""
        .LatestNews = "": .Location = "": .MaritalStatus = "": .MemberSince = "": .NickName = ""
        .Occupation = "": .RealName = "": .YahooID = ""
    End With
    If InStr(Data, "Member Since:") Then
        BackUP = Mid(Data, InStr(Data, "Member Since:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.MemberSince = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Last Updated:") Then
        BackUP = Mid(Data, InStr(Data, "Last Updated:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.LastUpdated = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "My Email") Then
        BackUP = Mid(Data, InStr(Data, "My Email"))
        If InStr(BackUP, "<i>Private</i>") Then
            NewProfile.Email = "Private"
            GoTo EndEmail
        ElseIf InStr(BackUP, "mailto:") Then
            LeftPos = InStr(BackUP, "mailto:") + 7
            RightPos = InStr(BackUP, Chr(34) & ">")
            NewProfile.Email = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
        Else
            NewProfile.Email = "Private"
        End If
EndEmail:
    End If
    DoEvents
    
    If InStr(Data, "Yahoo! ID:") Then
        BackUP = Mid(Data, InStr(Data, "Yahoo! ID:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.YahooID = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Real&nbsp;Name:") Then
        BackUP = Mid(Data, InStr(Data, "Real&nbsp;Name:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.RealName = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Nickname:") Then
        BackUP = Mid(Data, InStr(Data, "Nickname:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.NickName = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Location:") Then
        BackUP = Mid(Data, InStr(Data, "Location:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.Location = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Age:") Then
        BackUP = Mid(Data, InStr(Data, "Age:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.Age = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Marital&nbsp;Status:") Then
        BackUP = Mid(Data, InStr(Data, "Marital&nbsp;Status:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.MaritalStatus = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Gender:") Then
        BackUP = Mid(Data, InStr(Data, "Gender:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.Gender = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Occupation:") Then
        BackUP = Mid(Data, InStr(Data, "Occupation:"))
        LeftPos = InStr(BackUP, "<b>") + 3
        RightPos = InStr(BackUP, "</b>")
        NewProfile.Occupation = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Hobbies:") Then
        BackUP = Mid(Data, InStr(Data, "Hobbies:"))
        LeftPos = InStr(BackUP, "</font>") + 7
        RightPos = InStr(BackUP, "<br>")
        NewProfile.Hobbies = ConvertHTML(Trim(Mid(BackUP, LeftPos, RightPos - LeftPos)))
    End If
    DoEvents
    
    If InStr(Data, "Latest News:") Then
        BackUP = Mid(Data, InStr(Data, "Latest News:"))
        LeftPos = InStr(BackUP, "</font>") + 7
        RightPos = InStr(BackUP, "</td>")
        NewProfile.LatestNews = ConvertHTML(Trim(Mid(BackUP, LeftPos, RightPos - LeftPos)))
    End If
    DoEvents
    
    If InStr(Data, "Favorite Quote") Then
        BackUP = Mid(Data, InStr(Data, "Favorite Quote"))
        LeftPos = InStr(BackUP, "<i>") + 3
        RightPos = InStr(BackUP, "</i>")
        NewProfile.FavoriteQuote = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Home Page:") Then
        BackUP = Mid(Data, InStr(Data, "Home Page:"))
        If InStr(BackUP, "No home page specified") Then
            NewProfile.HomePage = "No home page specified"
            GoTo NextHomePage
        End If
        LeftPos = InStr(BackUP, "<a href=") + 9
        RightPos = InStr(BackUP, Chr(34) & ">")
        NewProfile.HomePage = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
NextHomePage:
    DoEvents
    
    If InStr(Data, "Cool Link 1:") Then
        BackUP = Mid(Data, InStr(Data, "Cool Link 1:"))
        LeftPos = InStr(BackUP, "<a href=") + 9
        RightPos = InStr(BackUP, Chr(34) & ">")
        NewProfile.CoolLink1 = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Cool Link 2:") Then
        BackUP = Mid(Data, InStr(Data, "Cool Link 2:"))
        LeftPos = InStr(BackUP, "<a href=") + 9
        RightPos = InStr(BackUP, Chr(34) & ">")
        NewProfile.CoolLink2 = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
    
    If InStr(Data, "Cool Link 3:") Then
        BackUP = Mid(Data, InStr(Data, "Cool Link 3:"))
        LeftPos = InStr(BackUP, "<a href=") + 9
        RightPos = InStr(BackUP, Chr(34) & ">")
        NewProfile.CoolLink3 = ConvertHTML(Mid(BackUP, LeftPos, RightPos - LeftPos))
    End If
    DoEvents
End Sub

Function ConvertHTML(TheString As String) As String
    ConvertHTML = TheString
    ConvertHTML = Replace(ConvertHTML, "&quot;", Chr(34))
    ConvertHTML = Replace(ConvertHTML, "&amp;", "&")
    ConvertHTML = Replace(ConvertHTML, "&lt;", "<")
    ConvertHTML = Replace(ConvertHTML, "&gt;", ">")
    ConvertHTML = Replace(ConvertHTML, "&nbsp;", " ")
    ConvertHTML = Replace(ConvertHTML, "&iexcl;", "¡")
    ConvertHTML = Replace(ConvertHTML, "&curren;", "¤")
    ConvertHTML = Replace(ConvertHTML, "&cent;", "¢")
    ConvertHTML = Replace(ConvertHTML, "&pound;", "£")
    ConvertHTML = Replace(ConvertHTML, "&yen;", "¥")
    ConvertHTML = Replace(ConvertHTML, "&brvbar;", "¦")
    ConvertHTML = Replace(ConvertHTML, "&sect;", "§")
    ConvertHTML = Replace(ConvertHTML, "&uml;", "¨")
    ConvertHTML = Replace(ConvertHTML, "&copy;", "©")
    ConvertHTML = Replace(ConvertHTML, "&ordf;", "ª")
    ConvertHTML = Replace(ConvertHTML, "&laquo;", "«")
    ConvertHTML = Replace(ConvertHTML, "&not;", "¬")
    ConvertHTML = Replace(ConvertHTML, "&shy;", "­")
    ConvertHTML = Replace(ConvertHTML, "&reg;", "®")
    ConvertHTML = Replace(ConvertHTML, "&trade;", "™")
    ConvertHTML = Replace(ConvertHTML, "&macr;", "¯")
    ConvertHTML = Replace(ConvertHTML, "&deg;", "°")
    ConvertHTML = Replace(ConvertHTML, "&plusmn;", "±")
    ConvertHTML = Replace(ConvertHTML, "&sup2;", "²")
    ConvertHTML = Replace(ConvertHTML, "&sup3;", "³")
    ConvertHTML = Replace(ConvertHTML, "&acute;", "´")
    ConvertHTML = Replace(ConvertHTML, "&micro;", "µ")
    ConvertHTML = Replace(ConvertHTML, "&para;", "¶")
    ConvertHTML = Replace(ConvertHTML, "&middot;", "·")
    ConvertHTML = Replace(ConvertHTML, "&cedil;", "¸")
    ConvertHTML = Replace(ConvertHTML, "&sup1;", "¹")
    ConvertHTML = Replace(ConvertHTML, "&ordm;", "º")
    ConvertHTML = Replace(ConvertHTML, "&raquo;", "»")
    ConvertHTML = Replace(ConvertHTML, "&frac14;", "¼")
    ConvertHTML = Replace(ConvertHTML, "&frac12;", "½")
    ConvertHTML = Replace(ConvertHTML, "&frac34;", "¾")
    ConvertHTML = Replace(ConvertHTML, "&iquest;", "¿")
    ConvertHTML = Replace(ConvertHTML, "&times;", "×")
    ConvertHTML = Replace(ConvertHTML, "&divide;", "÷")
    ConvertHTML = Replace(ConvertHTML, "&Agrave;", "À")
    ConvertHTML = Replace(ConvertHTML, "&Aacute;", "Á")
    ConvertHTML = Replace(ConvertHTML, "&Acirc;", "Â")
    ConvertHTML = Replace(ConvertHTML, "&Atilde;", "Ã")
    ConvertHTML = Replace(ConvertHTML, "&Auml;", "Ä")
    ConvertHTML = Replace(ConvertHTML, "&Aring;", "Å")
    ConvertHTML = Replace(ConvertHTML, "&AElig;", "Æ")
    ConvertHTML = Replace(ConvertHTML, "&Ccedil;", "Ç")
    ConvertHTML = Replace(ConvertHTML, "&Egrave;", "È")
    ConvertHTML = Replace(ConvertHTML, "&Eacute;", "É")
    ConvertHTML = Replace(ConvertHTML, "&Ecirc;", "Ê")
    ConvertHTML = Replace(ConvertHTML, "&Euml;", "Ë")
    ConvertHTML = Replace(ConvertHTML, "&Igrave;", "Ì")
    ConvertHTML = Replace(ConvertHTML, "&Iacute;", "Í")
    ConvertHTML = Replace(ConvertHTML, "&Icirc;", "Î")
    ConvertHTML = Replace(ConvertHTML, "&Iuml;", "Ï")
    ConvertHTML = Replace(ConvertHTML, "&ETH;", "Ð")
    ConvertHTML = Replace(ConvertHTML, "&Ntilde;", "Ñ")
    ConvertHTML = Replace(ConvertHTML, "&Ograve;", "Ò")
    ConvertHTML = Replace(ConvertHTML, "&Oacute;", "Ó")
    ConvertHTML = Replace(ConvertHTML, "&Ocirc;", "Ô")
    ConvertHTML = Replace(ConvertHTML, "&Otilde;", "Õ")
    ConvertHTML = Replace(ConvertHTML, "&Ouml;", "Ö")
    ConvertHTML = Replace(ConvertHTML, "&Oslash;", "Ø")
    ConvertHTML = Replace(ConvertHTML, "&Ugrave;", "Ù")
    ConvertHTML = Replace(ConvertHTML, "&Uacute;", "Ú")
    ConvertHTML = Replace(ConvertHTML, "&Ucirc;", "Û")
    ConvertHTML = Replace(ConvertHTML, "&UUml;", "Ü")
    ConvertHTML = Replace(ConvertHTML, "&Yacute;", "Ý")
    ConvertHTML = Replace(ConvertHTML, "&THORN;", "Þ")
    ConvertHTML = Replace(ConvertHTML, "&szlig;", "ß")
    ConvertHTML = Replace(ConvertHTML, "&agrave;", "à")
    ConvertHTML = Replace(ConvertHTML, "&aacute;", "á")
    ConvertHTML = Replace(ConvertHTML, "&acirc;", "â")
    ConvertHTML = Replace(ConvertHTML, "&atilde;", "ã")
    ConvertHTML = Replace(ConvertHTML, "&auml;", "ä")
    ConvertHTML = Replace(ConvertHTML, "&aring;", "å")
    ConvertHTML = Replace(ConvertHTML, "&aelig;", "æ")
    ConvertHTML = Replace(ConvertHTML, "&ccedil;", "ç")
    ConvertHTML = Replace(ConvertHTML, "&egrave;", "è")
    ConvertHTML = Replace(ConvertHTML, "&eacute;", "é")
    ConvertHTML = Replace(ConvertHTML, "&ecirc;", "ê")
    ConvertHTML = Replace(ConvertHTML, "&euml;", "ë")
    ConvertHTML = Replace(ConvertHTML, "&igrave;", "ì")
    ConvertHTML = Replace(ConvertHTML, "&iacute;", "í")
    ConvertHTML = Replace(ConvertHTML, "&icirc;", "î")
    ConvertHTML = Replace(ConvertHTML, "&iuml;", "ï")
    ConvertHTML = Replace(ConvertHTML, "&eth;", "ð")
    ConvertHTML = Replace(ConvertHTML, "&ntilde;", "ñ")
    ConvertHTML = Replace(ConvertHTML, "&ograve;", "ò")
    ConvertHTML = Replace(ConvertHTML, "&oacute;", "ó")
    ConvertHTML = Replace(ConvertHTML, "&ocirc;", "ô")
    ConvertHTML = Replace(ConvertHTML, "&otilde;", "õ")
    ConvertHTML = Replace(ConvertHTML, "&ouml;", "ö")
    ConvertHTML = Replace(ConvertHTML, "&oslash;", "ø")
    ConvertHTML = Replace(ConvertHTML, "&ugrave;", "ù")
    ConvertHTML = Replace(ConvertHTML, "&uacute;", "ú")
    ConvertHTML = Replace(ConvertHTML, "&ucirc;", "û")
    ConvertHTML = Replace(ConvertHTML, "&uuml;", "ü")
    ConvertHTML = Replace(ConvertHTML, "&yacute;", "ý")
    ConvertHTML = Replace(ConvertHTML, "&thorn;", "þ")
    ConvertHTML = Replace(ConvertHTML, "&yuml;", "ÿ")
End Function
