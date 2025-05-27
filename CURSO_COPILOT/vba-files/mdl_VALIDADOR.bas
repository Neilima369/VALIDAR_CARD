Attribute VB_Name ="mdl_VALIDADOR"

Public sub Verificar()
Dim cartaoNumero As String

cartaoNumero = "4556193141222516"
'''CALL FUNCTION CREDIT CARD VALIDATE
numeroCartao = mdl_VALIDADOR.Validar_Bandeiras_de_Cartao(cartaoNumero)
MsgBox numeroCartao, vbOKOnly, "VALIDAR CARTAO"

End sub

Public Function Validar_Bandeiras_de_Cartao(numeroCartao As String) As String
'''AUTOR ALUNO: CLAUDINEI LIMA - 25/05/2025 ÀS 19:30
'''E-MAIL: claudiflima@gmail.com

Dim chave As Variant
Dim regex As Object
Dim bandeiras As Object

Set regex = CreateObject("VBScript.RegExp")

Set bandeiras = CreateObject("Scripting.Dictionary")
    
bandeiras.Add "Visa", "^4[0-9]{12}(?:[0-9]{3})?$"
bandeiras.Add "Mastercard", "^5[1-5][0-9]{14}$"
bandeiras.Add "American Express", "^3[47][0-9]{13}$"
bandeiras.Add "Diners Club", "^3(?:0[0-5]|[68][0-9])[0-9]{11}$"
bandeiras.Add "Discover", "^6(?:011|5[0-9]{2})[0-9]{12}$"
bandeiras.Add "JCB", "^35[0-9]{14}$"
bandeiras.Add "EnRoute", "^(2014|2149)[0-9]{11}$"
bandeiras.Add "Voyager", "^8699[0-9]{1}[0-9]{10}$"
bandeiras.Add "HyperCard", "^6062[0-9]{12}$"
bandeiras.Add "Aura", "^50[0-9]{2}[0-9]{12}$"
    
Validar_Bandeiras_de_Cartao = "Bandeira desconhecida"
    
    For Each chave In bandeiras.Keys
        regex.Pattern = bandeiras(chave)
        If regex.Test(numeroCartao) Then
            Validar_Bandeiras_de_Cartao = chave
            Exit For
        End If
    Next chave
    Debug.Print Validar_Bandeiras_de_Cartao
End Function