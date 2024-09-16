Sub TransferirInformacoesLooping()

    Dim planilhaBanco As Worksheet
    Dim planilhaFormulario As Worksheet
    Set planilhaBanco = ThisWorkbook.Sheets("Sheet1")
    Set planilhaFormulario = ThisWorkbook.Sheets("formulario")

    
    Dim numeroLinha As Long
    numeroLinha = 2

    '########## PLANILHA DE ORIGEM DOS DADOS SHEET1 #########
    Dim mapeamento As Variant
    mapeamento = Array("G2", "H2", "L2", "M2", "U2", "I2", "J2", "N2", "Q2", "R2", "V2", "W2", "X2", "Y2", _
    "Z2", "AA2", "AB2", "AC2", "AD2", "AE2", "AF2", "AG2", "AH2", "AL2", "AJ2", "AK2", "AL2", "AM2", "AN2", "AO2", _
    "AP2", "AQ2", "AR2", "AS2", "AT2", "AU2", "AV2", "AW2", "BG2", "BH2", "BI2", "BJ2", "BK2", "BL2", "BO2", "BP2", _
    "BY2", "CB2", "CE2", "CF2", "CG2", "CH2", "CL2", "CM2", "CN2", "CO2", "CS2", "CT2", "CU2", "CV2", "CZ2", "DA2", _
    "DB2", "DC2", "DH2", "DI2", "DJ2", "DK2", "BN2", "BX2", "AZ2", "BC2")
    
    ' ######### PLANILHA DE DESTINO DOS DADOS FORMULARIO ###########
    Dim celulasDestino As Variant
    celulasDestino = Array("C8", "C9", "C11", "E11", "E21", "F9", "C10", "E12", "C17", "C18", "C22", "C23", "E23", "C24", _
    "C25", "E25", "C27", "E27", "C28", "E28", "C29", "E29", "C30", "E30", "C31", "E31", "C32", "E32", "C33", "E33", _
    "C34", "E34", "C35", "E35", "C36", "E36", "C37", "E37", "C41", "E41", "C43", "C44", "C45", "C46", "C62", "C63", _
    "C71", "E72", "C88", "E88", "C89", "E89", "C93", "E93", "C94", "E94", "C98", "E98", "C99", "E99", "C103", "E103", _
    "C104", "E104", "C110", "C111", "E111", "C112", "F48", "E69", "E38", "E39")
    

    Dim i As Integer
    For i = 0 To UBound(mapeamento)
       
        Dim celulaOrigem As Range
        Set celulaOrigem = planilhaBanco.Range(mapeamento(i))

        Dim celulaDestino As Range
        Set celulaDestino = planilhaFormulario.Range(celulasDestino(i))

        Dim valorOrigem As Variant
        valorOrigem = celulaOrigem.Value

        Dim valorAtualFormulario As Variant
        valorAtualFormulario = celulaDestino.Value

        celulaDestino.Value = valorAtualFormulario & " " & valorOrigem
    Next i
End Sub



Sub MasculinoOuFeminino()
    If ThisWorkbook.Sheets("Sheet1").Range("K2").Value = "Masculino" Then
        ThisWorkbook.Sheets("formulario").Range("E10").Value = "Sexo: ( x ) Masculino (  ) Feminino"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("K2").Value = "Feminino" Then
        ThisWorkbook.Sheets("formulario").Range("E10").Value = "Sexo: (  ) Masculino ( x ) Feminino"
   End If
End Sub


Sub EstadoCivil()
    
    If ThisWorkbook.Sheets("Sheet1").Range("O2").Value = "Solteiro" Then
        ThisWorkbook.Sheets("formulario").Range("C13").Value = "Estado civil: ( X ) Solteiro (  ) Casado (  ) Divorciado (  ) Viúvo (  ) União Estável (  ) Outros"
    
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("O2").Value = "Casado" Then
        ThisWorkbook.Sheets("formulario").Range("C13").Value = "Estado civil:  (  ) Solteiro ( X ) Casado (  ) Divorciado (  ) Viúvo (  ) União Estável (  ) Outros"
    
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("O2").Value = "Divorciado" Then
        ThisWorkbook.Sheets("formulario").Range("C13").Value = "Estado civil:  (  ) Solteiro (  ) Casado ( X ) Divorciado (  ) Viúvo (  ) União Estável (  ) Outros"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("O2").Value = "Viúvo" Then
        ThisWorkbook.Sheets("formulario").Range("C13").Value = "Estado civil:  (  ) Solteiro (  ) Casado (  ) Divorciado ( X ) Viúvo (  ) União Estável (  ) Outros"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("O2").Value = "União Estável" Then
        ThisWorkbook.Sheets("formulario").Range("C13").Value = "Estado civil:  (  ) Solteiro (  ) Casado ( X ) Divorciado (  ) Viúvo ( X ) União Estável (  ) Outros"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("O2").Value = "Outros" Then
        ThisWorkbook.Sheets("formulario").Range("C13").Value = "Estado civil:  (  ) Solteiro (  ) Casado ( X ) Divorciado (  ) Viúvo (  ) União Estável ( X ) Outros"
    End If
End Sub



Sub RacaCor()
    
    If ThisWorkbook.Sheets("Sheet1").Range("P2").Value = "Indígena" Then
        ThisWorkbook.Sheets("formulario").Range("C14").Value = "Raça e Cor:    ( X )Indígena    (  ) Branca   (  ) Negra   (  ) Amarela de origem japonesa, coreana etc."
    
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("P2").Value = "Branca" Then
        ThisWorkbook.Sheets("formulario").Range("C14").Value = "Raça e Cor:    (  )Indígena    ( X ) Branca   (  ) Negra   (  ) Amarela de origem japonesa, coreana etc."
    
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("P2").Value = "Negra" Then
        ThisWorkbook.Sheets("formulario").Range("C14").Value = "Raça e Cor:    (  )Indígena    (  ) Branca   ( X ) Negra   (  ) Amarela de origem japonesa, coreana etc."
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("P2").Value = "Amarela de origem japonesa, coreana etc." Then
        ThisWorkbook.Sheets("formulario").Range("C14").Value = "Raça e Cor:    (  )Indígena    (  ) Branca   (  ) Negra   ( X ) Amarela de origem japonesa, coreana etc."
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("P2").Value = "Parda (declarada como mulata, ou mestiça de negro com pessoa de outra cor ou raça)" Then
        ThisWorkbook.Sheets("formulario").Range("C15").Value = "( X )Parda (declarada como mulata, ou mestiça de negro com pessoa de outra cor ou raça)"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("P2").Value = "Não informado" Then
        ThisWorkbook.Sheets("formulario").Range("C16").Value = "( X ) Não informado"
    End If
    
End Sub


Sub PrimeiroEmprego()
   
    If ThisWorkbook.Sheets("Sheet1").Range("S2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C19").Value = "Primeiro emprego:  ( X ) Sim  (  ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("S2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C19").Value = "Primeiro emprego:  (  ) Sim  ( X ) Não"
    End If
End Sub


Sub ResidenteExterior()
   
    If ThisWorkbook.Sheets("Sheet1").Range("T2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C21").Value = "Residente no Exterior:  ( X ) Sim  (  ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("T2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C21").Value = "Residente no Exterior:  (  ) Sim  ( X ) Não"
    End If
End Sub



Sub Escolaridade()
    
    If ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "01 – Analfabeto" Then
        ThisWorkbook.Sheets("formulario").Range("C49").Value = "( X ) 01 – Analfabeto"
    
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "02 – Até a 4º série incompleta do ensino fundamental (antigo 1º grau ou primário)" Then
        ThisWorkbook.Sheets("formulario").Range("C50").Value = "( X ) 02 – Até a 4º série incompleta do ensino fundamental (antigo 1º grau ou primário)"
    
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "03 – 4º série completa do ensino fundamental (antigo 1º grau ou ginásio)" Then
        ThisWorkbook.Sheets("formulario").Range("C51").Value = "( X ) 03 – 4º série completa do ensino fundamental (antigo 1º grau ou ginásio)"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "04 – Da 5º a 8º série do ensino fundamental (antigo 1º grau ou ginásio)" Then
        ThisWorkbook.Sheets("formulario").Range("C52").Value = "( X ) 04 – Da 5º a 8º série do ensino fundamental (antigo 1º grau ou ginásio)"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "05 – Ensino fundamental completo (antigo 1º grau, primário ou ginásio)" Then
        ThisWorkbook.Sheets("formulario").Range("C53").Value = "( X ) 05 – Ensino fundamental completo (antigo 1º grau, primário ou ginásio)"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "06 – Ensino médio incompleto (antigo 2º grau, secundário ou colegial)" Then
        ThisWorkbook.Sheets("formulario").Range("C54").Value = "( X ) 06 – Ensino médio incompleto (antigo 2º grau, secundário ou colegial)"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "07 – Ensino médio completo (antigo 2º grau, secundário ou colegial)" Then
        ThisWorkbook.Sheets("formulario").Range("C55").Value = "( X ) 07 – Ensino médio completo (antigo 2º grau, secundário ou colegial)"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "08 – Educação Superior incompleta" Then
        ThisWorkbook.Sheets("formulario").Range("C56").Value = "( X ) 08 – Educação Superior incompleta"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "09 – Educação Superior completa" Then
        ThisWorkbook.Sheets("formulario").Range("C57").Value = "( X ) 09 – Educação Superior completa"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "10 – Pós Graduação" Then
        ThisWorkbook.Sheets("formulario").Range("C58").Value = "( X ) 10 – Pós Graduação"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "11 – Mestrado" Then
        ThisWorkbook.Sheets("formulario").Range("C59").Value = "( X ) 11 – Mestrado"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BM2").Value = "12 - Doutorado" Then
        ThisWorkbook.Sheets("formulario").Range("C60").Value = "( X ) 12 - Doutorado"

    End If
    
End Sub



Sub Dependentes()
    
    If ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "Cônjuge ou companheiro (a) com o (a) qual tenha filho ou viva a mais de 5 (cinco) anos;" Then
        ThisWorkbook.Sheets("formulario").Range("C76").Value = "( X ) 01 – Cônjuge ou companheiro (a) com o (a) qual tenha filho ou viva a mais de 5 (cinco) anos;"
    
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "02 – Filho (a) ou enteado (a) até 21 (vinte e um) anos;" Then
        ThisWorkbook.Sheets("formulario").Range("C77").Value = "( X ) 02 – Filho (a) ou enteado (a) até 21 (vinte e um) anos;"
    
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "03 – Filho (a) ou enteado (a) universitário (a) ou cursando escola técnica de 2º grau, até 24 (vinte e quatro) anos;" Then
        ThisWorkbook.Sheets("formulario").Range("C78").Value = "( X ) 03 – Filho (a) ou enteado (a) universitário (a) ou cursando escola técnica de 2º grau, até 24 (vinte e quatro) anos;"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "04 – Filho (a) ou enteado (a) em qualquer idade, quando incapacitado física e/ou mentalmente para o trabalho;" Then
        ThisWorkbook.Sheets("formulario").Range("C79").Value = "( X ) 04 – Filho (a) ou enteado (a) em qualquer idade, quando incapacitado física e/ou mentalmente para o trabalho;"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "05 – Irmão (a), neto (a) ou bisneto (a) sem arrimo dos pais, do (a) qual detenha a guarda judicial, até 21 (vinte um) anos;" Then
        ThisWorkbook.Sheets("formulario").Range("C80").Value = "( X ) 05 – Irmão (a), neto (a) ou bisneto (a) sem arrimo dos pais, do (a) qual detenha a guarda judicial, até 21 (vinte um) anos;"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "06 – Irmão (a), neto (a) ou bisneto (a) sem arrimo dos pais, com idade até 24 anos, se ainda estiver cursando estabelecimento de nível superior ou escola técnica de 2º grau, desde que tenha detido sua guarda judicial até os 21 anos;" Then
        ThisWorkbook.Sheets("formulario").Range("C81").Value = "( X ) 06 – Irmão (a), neto (a) ou bisneto (a) sem arrimo dos pais,com idade até 24 anos, se ainda estiver cursando estabelecimento de nível superior"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "07 - Irmão (a), neto (a) ou bisneto (a) sem arrimo dos pais, do (a) qual detenha a guarda judicial, em qualquer idade, quando incapacitado física e/ou mentalmente para o trabalho;" Then
        ThisWorkbook.Sheets("formulario").Range("C83").Value = "( X ) 07 - Irmão (a), neto (a) ou bisneto (a) sem arrimo dos pais, do (a) qual detenha a guarda judicial, em qualquer idade, quando incapacitado física"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "08 – Pais,avós e bisavós" Then
        ThisWorkbook.Sheets("formulario").Range("C85").Value = "( X ) 08 – Pais,avós e bisavós"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "09- Menor pobre, até 21 (vinte e um anos), que crie e eduque e do qual detenha a guarda judicial;" Then
        ThisWorkbook.Sheets("formulario").Range("C86").Value = "( X ) 09- Menor pobre, até 21 (vinte e um anos), que crie e eduque e do qual detenha a guarda judicial;"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CD2").Value = "10 – A pessoa absolutamente incapaz, da qual seja tutor ou curador." Then
        ThisWorkbook.Sheets("formulario").Range("C87").Value = "10 – A pessoa absolutamente incapaz, da qual seja tutor ou curador."

    End If
    
End Sub

Sub trabalhadorEstrangeiro()
   
    If ThisWorkbook.Sheets("Sheet1").Range("BQ2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C64").Value = "Condição de casado com brasileiros em caso de trabalhador estrangeiro: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BQ2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C64").Value = "Condição de casado com brasileiros em caso de trabalhador estrangeiro: (  ) Sim ( X ) Não"
    End If
    
    ' SE O TRABALHADOR ESTRANGEIRO TEM FILHOS COM BRASILEIRO
        If ThisWorkbook.Sheets("Sheet1").Range("BR2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C65").Value = "Se o trabalhador estrangeiro tem filhos com brasileiro: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BR2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C65").Value = "Se o trabalhador estrangeiro tem filhos com brasileiro: (  ) Sim ( X ) Não"
    End If
End Sub



Sub PCD()
   
    If ThisWorkbook.Sheets("Sheet1").Range("BS2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C67").Value = "Deficiência física: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BS2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C67").Value = "Deficiência física: (  ) Sim ( X ) Não"
    End If
    
    ' DEFICIENCIA MENTAL
        If ThisWorkbook.Sheets("Sheet1").Range("BT2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("E67").Value = "Deficiência Mental: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BT2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("E67").Value = "Deficiência Mental: (  ) Sim ( X ) Não"
    End If

    ' DEFICIENCIA VISUAL
    If ThisWorkbook.Sheets("Sheet1").Range("BU2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C68").Value = "Deficiência visual: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BU2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C68").Value = "Deficiência visual: (  ) Sim ( X ) Não"
    End If
    
    ' DEFICIENCIA INTELECTUAL
    If ThisWorkbook.Sheets("Sheet1").Range("BV2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("E68").Value = "Deficiência Intelectual: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BV2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("E68").Value = "Deficiência Intelectual: (  ) Sim ( X ) Não"
    End If
    
    ' DEFICIENCIA AUDITIVA
    If ThisWorkbook.Sheets("Sheet1").Range("BW2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C69").Value = "Deficiência auditiva: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("BW2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C69").Value = "Deficiência auditiva: (  ) Sim ( X ) Não"
    End If
End Sub



Sub ContaBancaria()

    If ThisWorkbook.Sheets("Sheet1").Range("CC2").Value = "Conta Corrente" Then
        ThisWorkbook.Sheets("formulario").Range("C73").Value = "Tipo da Conta: ( X ) Conta Corrente    (  ) Conta poupança   (  ) Conta Salário"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CC2").Value = "Conta poupança" Then
        ThisWorkbook.Sheets("formulario").Range("C73").Value = "Tipo da Conta: (  ) Conta Corrente    ( X ) Conta poupança   (  ) Conta Salário"
        
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CC2").Value = "Conta Salário" Then
        ThisWorkbook.Sheets("formulario").Range("C73").Value = "Tipo da Conta: (  ) Conta Corrente    (  ) Conta poupança   ( X ) Conta Salário"
    End If
    
End Sub



Sub IRRF()
   ' DEPENDENTE 1
    If ThisWorkbook.Sheets("Sheet1").Range("CI2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C90").Value = "Dependentes para fins de IRRF: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CI2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C90").Value = "Dependentes para fins de IRRF: (  ) Sim ( X ) Não"
    End If
    
    If ThisWorkbook.Sheets("Sheet1").Range("CJ2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C91").Value = "Dependentes para fins de Salário-Família: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CJ2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C91").Value = "Dependentes para fins de Salário-Família: (  ) Sim ( X ) Não"
    End If

    If ThisWorkbook.Sheets("Sheet1").Range("CK2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C92").Value = "Há incapacidade física ou mental para o trabalho: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CK2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C92").Value = "Há incapacidade física ou mental para o trabalho: (  ) Sim ( X ) Não"
    End If
    
    
    'DEPENDENTE 2
    If ThisWorkbook.Sheets("Sheet1").Range("CP2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C95").Value = "Dependentes para fins de IRRF: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CP2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C95").Value = "Dependentes para fins de IRRF: (  ) Sim ( X ) Não"
    End If
    
    If ThisWorkbook.Sheets("Sheet1").Range("CQ2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C96").Value = "Dependentes para fins de Salário-Família: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CQ2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C96").Value = "Dependentes para fins de Salário-Família: (  ) Sim ( X ) Não"
    End If

    If ThisWorkbook.Sheets("Sheet1").Range("CR2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C97").Value = "Há incapacidade física ou mental para o trabalho: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CR2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C97").Value = "Há incapacidade física ou mental para o trabalho: (  ) Sim ( X ) Não"
    End If
    
    
    'DEPENDENTE 3
    If ThisWorkbook.Sheets("Sheet1").Range("CW2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C100").Value = "Dependentes para fins de IRRF: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CW2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C100").Value = "Dependentes para fins de IRRF: (  ) Sim ( X ) Não"
    End If
    
    If ThisWorkbook.Sheets("Sheet1").Range("CX2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C101").Value = "Dependentes para fins de Salário-Família: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CX2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C101").Value = "Dependentes para fins de Salário-Família: (  ) Sim ( X ) Não"
    End If

    If ThisWorkbook.Sheets("Sheet1").Range("CY2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C102").Value = "Há incapacidade física ou mental para o trabalho: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("CY2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C102").Value = "Há incapacidade física ou mental para o trabalho: (  ) Sim ( X ) Não"
    End If
    
    
    'DEPENDENTE 4
    If ThisWorkbook.Sheets("Sheet1").Range("DD2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C105").Value = "Dependentes para fins de IRRF: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("DD2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C105").Value = "Dependentes para fins de IRRF: (  ) Sim ( X ) Não"
    End If
    
    If ThisWorkbook.Sheets("Sheet1").Range("DE2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C106").Value = "Dependentes para fins de Salário-Família: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("DE2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C106").Value = "Dependentes para fins de Salário-Família: (  ) Sim ( X ) Não"
    End If

    If ThisWorkbook.Sheets("Sheet1").Range("DF2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C107").Value = "Há incapacidade física ou mental para o trabalho: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("DF2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C107").Value = "Há incapacidade física ou mental para o trabalho: (  ) Sim ( X ) Não"
    End If
    
    
    ' MULTIPLOS VINCULOS TRABALHISTAS
    If ThisWorkbook.Sheets("Sheet1").Range("DG2").Value = "Sim" Then
        ThisWorkbook.Sheets("formulario").Range("C109").Value = "Trabalha registrado em outra empresa: ( X ) Sim ( ) Não"
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("DG2").Value = "Não" Then
        ThisWorkbook.Sheets("formulario").Range("C109").Value = "Trabalha registrado em outra empresa: (  ) Sim ( X ) Não"
    End If
    End Sub
    



Sub ConcatenarCelulas()
    Dim planilhaBanco As Worksheet
    Dim planilhaFormulario As Worksheet
    Dim numeroCNH As String
    Dim dataExpedicao As String
    Dim ufCNH As String
    
    ' Definir planilhas de origem e destino
    Set planilhaBanco = ThisWorkbook.Sheets("Sheet1")
    Set planilhaFormulario = ThisWorkbook.Sheets("formulario")
    
    
    ' Obter número da CNH, data de expediï¿½ï¿½o e UF da CNH da planilha "Sheet1"
    numeroCNH = planilhaBanco.Range("AX2").Value
    categoria = planilhaBanco.Range("AY2").Value
    dataExpedicao = planilhaBanco.Range("BA2").Value
    ufCNH = planilhaBanco.Range("BB2").Value
    
    ' Titulo eleitoral
    zonaEleitoral = planilhaBanco.Range("BE2").Value
    secao = planilhaBanco.Range("BF2").Value
    
    'Codigo do banco]
    codBanco = planilhaBanco.Range("BZ2").Value
    codAgencia = planilhaBanco.Range("CA2").Value
    
    ' Preencher as celulas C8 e C39 na planilha "formulario" com as informações
    With planilhaFormulario
        .Range("C38").Value = "Número CNH: " & numeroCNH & " -                              45 - Categoria: " & categoria
        .Range("C39").Value = "Data de Expedição CNH: " & dataExpedicao & "                48 - UF da CNH: " & ufCNH
        .Range("E40").Value = "Zona: " & zonaEleitoral & "           Seção: " & secao
        
        .Range("C72").Value = "Código do banco: " & codBanco & "                72 - Código da agência: " & codAgencia
    End With
End Sub






Sub Preencher()

    TransferirInformacoesLooping
    
    MasculinoOuFeminino
    
    EstadoCivil
    
    RacaCor
    
    PrimeiroEmprego

    Escolaridade
    
    trabalhadorEstrangeiro
    
    PCD
    
    ContaBancaria
    
    IRRF
    
    Dependentes
    
    ConcatenarCelulas
End Sub
