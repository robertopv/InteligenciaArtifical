Attribute VB_Name = "JsonConverter"
Option Explicit
Public Function ObtenDados(JSON, Chave As String) As String
Dim i   As Integer
Dim ret As String

ret = Mid(JSON, InStr(JSON, Chave) + Len(Chave) + 4, Len(JSON))
ret = Mid(ret, InStr(JSON, """:""") + 4, Len(ret))

ret = Mid(ret, 1, InStr(ret, """,""") - 1)
ObtenDados = ret

End Function

