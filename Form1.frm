VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPTZINHO"
   ClientHeight    =   7110
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   11655
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5280
      Width           =   11655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "enviar"
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sc As New ScriptControl


Private Sub Command1_Click()
        Dim XMLHTTP As Object
        Dim myUrl As String
        Dim body As String
        Dim ret As String
        
        Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
        
        ' URL para enviar o POST
        myUrl = "https://api.openai.com/v1/completions"
        
        ' Token Bearer que será enviado no header
        myToken = "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"
        
        ' JSON que será enviado
            body = body & " {  "
            body = body & " ""model"": ""text-davinci-003"", "
            body = body & " ""prompt"": """ & Replace(Text1.Text, vbCrLf, "") & """, "
            body = body & " ""max_tokens"": 2000, "
            body = body & " ""temperature"": 0 "
            body = body & "}"
        
Text.Text = Text.Text & "PERGUNTA:" & vbNewLine
Text.Text = Text.Text & "-" & Text1.Text & vbNewLine & vbNewLine

Text1.Text = ""
       

' Envia o POST com o token Bearer no header
XMLHTTP.Open "POST", myUrl, False
XMLHTTP.setRequestHeader "Content-Type", "application/json"
XMLHTTP.setRequestHeader "Authorization", "Bearer " & myToken
XMLHTTP.send body
ret = XMLHTTP.responseText


'---------------------------------

    ' Define o JSON
    Dim JSON As String
    JSON = Trim(ObtenDados(ret, "choices"))
    

'------------------------------


        ' Exibe a resposta do servidor
        'MsgBox JSON
Text.Text = Text.Text & "RESPOSTA:" & vbNewLine
Text.Text = Text.Text & "-" & Replace(Mid(JSON, 5, Len(JSON)), "\n", vbNewLine) & vbNewLine
Text.Text = Text.Text & "_________________________________________________________________________________________________________________________" & vbNewLine & vbNewLine
Text.SelStart = Len(Text.Text)
Text.SelLength = 0

End Sub


Private Sub Text_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Text.SelStart = Len(Text.Text)
        Text.SelLength = 0
    End If
End Sub

Private Sub Text_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Text.SelStart = Len(Text.Text)
        Text.SelLength = 0
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
            Command1_Click
            Text1.SetFocus
            KeyAscii = 0
    End If
End Sub

