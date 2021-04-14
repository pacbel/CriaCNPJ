Attribute VB_Name = "basPrincipal"
Option Explicit

Public Function ValidarCNPJ_CPF(ByVal pValor As String) As Boolean

          Dim CPFComp As String
          Dim CGCComp As String
          Dim CPF1 As String
          Dim CGC1 As String
          Dim soma As Integer
          Dim Num As Integer
          Dim C As Integer
          Dim digito1 As String
          Dim digito2 As String

1     If Len(Trim(pValor)) > 11 Then
2       CGCComp = ZerosAEsquerda(pValor, 15)
3       CGC1 = Mid(CGCComp, 1, 13)
4       soma = 0
5       Num = 5
6       For C = 1 To 13
7           soma = soma + Val(Mid(CGC1, C, 1)) * Num
8           Num = Num + 1
9           Num = IIf(Num = 10, 2, Num)
10      Next
11      digito1 = IIf((soma Mod 11) = 10, "0", Mid(soma Mod 11, 1, 1))
12      CGC1 = CGC1 + digito1
13      soma = 0
14      Num = 4
15      For C = 1 To 14
16          soma = soma + Val(Mid(CGC1, C, 1)) * Num
17          Num = Num + 1
18          Num = IIf(Num = 10, 2, Num)
19      Next
20      digito2 = IIf((soma Mod 11) = 10, "0", Mid(soma Mod 11, 1, 1))
521      ValidarCNPJ_CPF = IIf(Mid(CGCComp, 14, 2) = digito1 + digito2, True, False)
    frmPrincipal.lblDigito.Caption = digito1 & digito2
22    Else
23      CPFComp = ZerosAEsquerda(pValor, 11)
24      CPF1 = Mid(CPFComp, 1, 9)
25      soma = 0
26      For C = 1 To 9
27          soma = soma + Val(Mid(CPF1, C, 1)) * C
28      Next
29      digito1 = IIf((soma Mod 11) = 10, "0", Mid((soma Mod 11), 1, 1))
30      CPF1 = Mid(CPFComp, 2, 8) + digito1
31      soma = 0
32      For C = 1 To 9
33          soma = soma + Val(Mid(CPF1, C, 1)) * C
34      Next
35      digito2 = IIf((soma Mod 11) = 10, "0", Mid(soma Mod 11, 1, 1))
36      ValidarCNPJ_CPF = IIf(Mid(CPFComp, 10, 2) = digito1 + digito2, True, False)
        frmPrincipal.lblDigito.Caption = digito1 & digito2
37    End If
End Function


Public Function ZerosAEsquerda(Texto As String, numcasas As Integer)
1     If Len(Texto) > numcasas Then
2       Texto = Right(Texto, numcasas)
3     End If
4     ZerosAEsquerda = String(numcasas - Len(Trim(Texto)), "0") + Trim(Texto)
End Function

