Attribute VB_Name = "Módulo1"
'Cálculo del diámetro equivalente del elemento de relleno para la determinación
'de la retención de líquido en torres rellenas [Tabla 6.5 Treybal 2ª Ed.]
Function diámetro_equivalente_Raschig_cerámica(diámetro_nominal)
Select Case diámetro_nominal    'en milímetros
   Case 13
   diámetro_equivalente_Raschig_cerámica = 0.01774
   Case 25
   diámetro_equivalente_Raschig_cerámica = 0.0356
   Case 38
   diámetro_equivalente_Raschig_cerámica = 0.053
   Case 50
   diámetro_equivalente_Raschig_cerámica = 0.0725
End Select
End Function

Function diámetro_equivalente_Raschig_carbón(diámetro_nominal)
Select Case diámetro_nominal    'en milímetros
   Case 25
   diámetro_equivalente_Raschig_carbón = 0.01301
   Case 38
   diámetro_equivalente_Raschig_ carbón = 0.0543
   Case 50
   diámetro_equivalente_Raschig_ carbón = 0.0716
End Select
End Function

Function diámetro_equivalente_Berl_cerámica(diámetro_nominal)
Select Case diámetro_nominal    'en milímetros
   Case 13
   diámetro_equivalente_Berl_cerámica = 0.01622
   Case 25
   diámetro_equivalente_Berl_cerámica = 0.032
   Case 38
   diámetro_equivalente_Berl_cerámica = 0.0472
End Select
End Function

'Programación alternativa para el anterior
Function diámetro_equivalente(diámetro_nominal, relleno, material)
Select Case relleno

Case "Anillos_Raschig"
            Select Case material
      Case "cerámica"
                Select Case diámetro_nominal
                Case 13
                diámetro_equivalente = 0.01774
Case 25
                diámetro_equivalente = 0.0356
Case 38
                diámetro_equivalente = 0.053
Case 50
                diámetro_equivalente = 0.0725
                End Select
Case "carbón"
                Select Case diámetro_nominal
                Case 25
                diámetro_equivalente = 0.01301
Case 38
                diámetro_equivalente = 0.0543
Case 50
                diámetro_equivalente = 0.0716
                End Select
    End Select

Case "Sillas_Berl"

          Select Case material
Case "carbón"
                Select Case diámetro_nominal
                Case 13
                diámetro_equivalente = 0.01622
Case 25
                diámetro_equivalente = 0.032
Case 38
                diámetro_equivalente = 0.0472
                End Select
    End Select
            
End Select

End Function
  




'Cálculo del coeficiente de transferencia de materia en la fase líquida para
'absorción de gases en torres rellenas con anillos Raschig y sillas de Berl
'Ec. 18-57 Perry 6ª Ed., Ec. 6.72 Treybal 2ª Ed. [Schulmann et al. AIChE J. 1, 253]
Function k_líquido_x_Raschig_cerámica(diámetro_relleno_in, flujo_másico_líquido, µ_líquido, Sc_líquido, difusividad_líquido, densidad_líquido, masa_molar_líquido)
'En el SI este coeficiente de transferencia tiene unidades de kmol/m2/s
Select Case diámetro_relleno_in
    Case 0.5
    k_líquido_x_Raschig_cerámica = 25.1 * (0.01774 * flujo_másico_líquido / µ_líquido) ^ 0.45 * Sc_líquido ^ 0.5 * difusividad_líquido * densidad_líquido / 0.01774 / masa_molar_líquido
    Case 1
    k_líquido_x_Raschig_cerámica = 25.1 * (0.0356 * flujo_másico_líquido / µ_líquido) ^ 0.45 * Sc_líquido ^ 0.5 * difusividad_líquido * densidad_líquido / 0.0356 / masa_molar_líquido
    Case 1.5
    k_líquido_x_Raschig_cerámica = 25.1 * (0.053 * flujo_másico_líquido / µ_líquido) ^ 0.45 * Sc_líquido ^ 0.5 * difusividad_líquido * densidad_líquido / 0.053 / masa_molar_líquido
    Case 2
    k_líquido_x_Raschig_cerámica = 25.1 * (0.0725 * flujo_másico_líquido / µ_líquido) ^ 0.45 * Sc_líquido ^ 0.5 * difusividad_líquido * densidad_líquido / 0.0725 / masa_molar_líquido
End Select
End Function

'Cálculo del coeficiente de transferencia de materia en la fase gaseosa para
'absorción de gases en torres rellenas con anillos Raschig y sillas de Berl
'Ec.  Perry 6ª Ed., Ec. 6.70 Treybal 2ª Ed. [Schulmann et al. AIChE J. 1, 253]
Function k_gas_y_Raschig_cerámica(diámetro_relleno_in, flujo_másico_gas, µ_gas, Sc_gas, esp_vacío_op, masa_molar_gas)
'En el SI este coeficiente de transferencia tiene unidades de kmol/m2/s
Select Case diámetro_relleno_in
    Case 0.5
    k_gas_y_Raschig_cerámica = 1.195 * (0.01744 * flujo_másico_gas / µ_gas / (1 - esp_vacío_op)) ^ (-0.36) * Sc_gas ^ (-2 / 3) * flujo_másico_gas / masa_molar_gas
    Case 1
    k_gas_y_Raschig_cerámica = 1.195 * (0.03556 * flujo_másico_gas / µ_gas / (1 - esp_vacío_op)) ^ (-0.36) * Sc_gas ^ (-2 / 3) * flujo_másico_gas / masa_molar_gas
    Case 1.5
    k_gas_y_Raschig_cerámica = 1.195 * (0.053 * flujo_másico_gas / µ_gas / (1 - esp_vacío_op)) ^ (-0.36) * Sc_gas ^ (-2 / 3) * flujo_másico_gas / masa_molar_gas
    Case 2
    k_gas_y_Raschig_cerámica = 1.195 * (0.0725 * flujo_másico_gas / µ_gas / (1 - esp_vacío_op)) ^ (-0.36) * Sc_gas ^ (-2 / 3) * flujo_másico_gas / masa_molar_gas
End Select
End Function

'Cálculo del número de Schmidt
Function Schmidt(difusividad, µ, densidad)
Schmidt = µ / difusividad / densidad
End Function

'Cálculo del número de Reynolds
Function Reynolds(densidad, velocidad, diámetro, µ)
Reynolds = densidad * velocidad * diámetro / µ
End Function

'Viscosidad de una mezcla gaseosa binaria, para hidrocarburos y no hidrocarburos
'a bajas presiones debajo de T reducida de 0,6. Ec. 2-100 Perry 7ª ed.
Function Viscosidad_mezcla_gaseosa_binaria(fracción_mol_1, viscosidad_1, viscosidad_2, masa_molar_1, masa_molar_2)
Q12 = (1 + ((viscosidad_1 / viscosidad_2) ^ 0.5 + (masa_molar_2 / masa_molar_1) ^ 0.25) ^ 2) / (8 * (1 + masa_molar_1 / masa_molar_2)) ^ 0.5
Q21 = (1 + ((viscosidad_2 / viscosidad_1) ^ 0.5 + (masa_molar_1 / masa_molar_2) ^ 0.25) ^ 2) / (8 * (1 + masa_molar_2 / masa_molar_1)) ^ 0.5
Viscosidad_mezcla_gaseosa_binaria = viscosidad_1 / (1 + Q12 * (1 - fracción_mol_1) / fracción_mol_1) + viscosidad_2 / (1 + Q21 * fracción_mol_1 / (1 - fracción_mol_1))
End Function

'Viscosidad de una mezcla líquida binaria, para mezcla de no hidrocarburos
'Ec. 2-120 Perry 7ma. ed.
Function Viscosidad_mezcla_líquida_binaria_noHC(fracción_mol_1, viscosidad_1, viscosidad_2)
Viscosidad_mezcla_líquida_binaria_noHC = Exp(fracción_mol_1 * Log(viscosidad_1) + (1 - fracción_mol_1) * Log(viscosidad_2))
End Function

'Viscosidad de una mezcla líquida binaria, para mezcla de hidrocarburos
'Ec. 2-119 Perry 7ma. ed.
Function Viscosidad_mezcla_líquida_binaria_HC(fracción_mol_1, viscosidad_1, viscosidad_2)
Viscosidad_mezcla_líquida_binaria_HC = (fracción_mol_1 * (viscosidad_1) ^ (1 / 3) + (1 - fracción_mol_1) * (viscosidad_2) ^ (1 / 3)) ^ 3
End Function

'Difusividad de una mezcla gaseosa binaria, para aire y hidrocarburo/no hidrocarburo
'Ec. 2-152 Perry 7ma. ed.
Function Difusividad_mezcla_gaseosa_binaria(temperatura_abs, masa_molar_1, masa_molar_2, presión_abs, contribución_grupo_1, contribución_grupo_2)
Difusividad_mezcla_gaseosa_binaria = 0.01013 * temperatura_abs ^ (1.75) * (1 / masa_molar_1 + 1 / masa_molar_2) ^ 0.5 / presión_abs / (contribución_grupo_1 ^ (1 / 3) + contribución_grupo_2 ^ (1 / 3)) ^ 2
End Function

'Cálculo del factor de corrección H para anillos Raschig de cerámica
Function Factor_corrección_H_Raschig_cerámica(flujo_másico_líquido, µ_líquido, densidad_líquido, tensión_superficial_líquido)
If (µ_líquido < 0.012) Then
Factor_corrección_H_Raschig_cerámica = 975.7 * flujo_másico_líquido ^ 0.57 * µ_líquido ^ 0.13 / densidad_líquido ^ 0.84 / (2.024 * flujo_másico_líquido ^ 0.43 - 1) * (tensión_superficial_líquido / 0.073) ^ (0.1737 - 0.262 * Log(flujo_másico_líquido) / 2.30259)
Else
Factor_corrección_H_Raschig_cerámica = 2168 * flujo_másico_líquido ^ 0.57 * µ_líquido ^ 0.31 / densidad_líquido ^ 0.84 / (2.024 * flujo_másico_líquido ^ 0.43 - 1) * (tensión_superficial_líquido / 0.073) ^ (0.1737 - 0.262 * Log(flujo_másico_líquido) / 2.30259)
End If
End Function

'Cálculo del factor de corrección H para anillos Raschig de carbón
Function Factor_corrección_H_Raschig_carbón(flujo_másico_líquido, µ_líquido, densidad_líquido, tensión_superficial_líquido)
If (µ_líquido < 0.012) Then
Factor_corrección_H_Raschig_carbón = 407.9 * flujo_másico_líquido ^ 0.57 * µ_líquido ^ 0.13 / densidad_líquido ^ 0.84 / (1.393 * flujo_másico_líquido ^ 0.315 - 1) * (tensión_superficial_líquido / 0.073) ^ (0.1737 - 0.262 * Log(flujo_másico_líquido) / 2.30259)
Else
Factor_corrección_H_Raschig_carbón = 901 * flujo_másico_líquido ^ 0.57 * µ_líquido ^ 0.31 / densidad_líquido ^ 0.84 / (1.393 * flujo_másico_líquido ^ 0.315 - 1) * (tensión_superficial_líquido / 0.073) ^ (0.1737 - 0.262 * Log(flujo_másico_líquido) / 2.30259)
End If
End Function

'Cálculo del factor de corrección H para sillas de Berl de cerámica
Function Factor_corrección_H_Berl_cerámica(flujo_másico_líquido, µ_líquido, densidad_líquido, tensión_superficial_líquido)
If (µ_líquido < 0.02) Then
Factor_corrección_H_Berl_cerámica = 1404 * flujo_másico_líquido ^ 0.57 * µ_líquido ^ 0.13 / densidad_líquido ^ 0.84 / (3.24 * flujo_másico_líquido ^ 0.413 - 1) * (tensión_superficial_líquido / 0.073) ^ (0.2817 - 0.262 * Log(flujo_másico_líquido) / 2.30259)
Else
Factor_corrección_H_Berl_cerámica = 2830 * flujo_másico_líquido ^ 0.57 * µ_líquido ^ 0.31 / densidad_líquido ^ 0.84 / (3.24 * flujo_másico_líquido ^ 0.413 - 1) * (tensión_superficial_líquido / 0.073) ^ (0.28177 - 0.262 * Log(flujo_másico_líquido) / 2.30259)
End If
End Function

'Cálculo del área interfacial para anillos Raschig en agua
Function Área_interfacial_agua_Raschig(flujo_másico_gas, densidad_gas, flujo_másico_líquido, diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
        If flujo_másico_líquido >= 0.68 And flujo_másico_líquido < 2 Then
        Área_interfacial_agua_Raschig = 28.01 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.2323 * flujo_másico_líquido - 0.3) * flujo_másico_líquido ^ -1.04
        ElseIf flujo_másico_líquido >= 2 And flujo_másico_líquido < 6.1 Then
        Área_interfacial_agua_Raschig = 14.69 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.01114 * flujo_másico_líquido + 0.148) * flujo_másico_líquido ^ -0.111
        End If
    Case 1
        If flujo_másico_líquido >= 0.68 And flujo_másico_líquido < 2 Then
        Área_interfacial_agua_Raschig = 34.42 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0) * flujo_másico_líquido ^ 0.552
        ElseIf flujo_másico_líquido >= 2 And flujo_másico_líquido < 6.1 Then
        Área_interfacial_agua_Raschig = 68.2 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.0389 * flujo_másico_líquido - 0.0793) * flujo_másico_líquido ^ -0.47
        End If
    Case 1.5
        If flujo_másico_líquido >= 0.68 And flujo_másico_líquido < 2 Then
        Área_interfacial_agua_Raschig = 36.5 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.0498 * flujo_másico_líquido - 0.1013) * flujo_másico_líquido ^ 0.274
        ElseIf flujo_másico_líquido >= 2 And flujo_másico_líquido < 6.1 Then
        Área_interfacial_agua_Raschig = 40.11 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.01091 * flujo_másico_líquido - 0.022) * flujo_másico_líquido ^ 0.14
        End If
    Case 2
        If flujo_másico_líquido >= 0.68 And flujo_másico_líquido < 2 Then
        Área_interfacial_agua_Raschig = 31.52 * flujo_másico_líquido ^ 0.481
        ElseIf flujo_másico_líquido >= 2 And flujo_másico_líquido < 6.1 Then
        Área_interfacial_agua_Raschig = 34.03 * flujo_másico_líquido ^ 0.362
        End If
End Select
End Function

'Cálculo del área interfacial para sillas de Berl en agua
Function Área_interfacial_agua_Berl(flujo_másico_gas, densidad_gas, flujo_másico_líquido, diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
        If flujo_másico_líquido >= 0.68 And flujo_másico_líquido < 2 Then
        Área_interfacial_agua_Berl = 16.28 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.0529) * flujo_másico_líquido ^ 0.761
        ElseIf flujo_másico_líquido >= 2 And flujo_másico_líquido < 6.1 Then
        Área_interfacial_agua_Berl = 25.61 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.01529) * flujo_másico_líquido ^ 0.17
        End If
    Case 1
        If flujo_másico_líquido >= 0.68 And flujo_másico_líquido < 2 Then
        Área_interfacial_agua_Berl = 52.14 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.0506 * flujo_másico_líquido - 0.1029)
        ElseIf flujo_másico_líquido >= 2 And flujo_másico_líquido < 6.1 Then
        Área_interfacial_agua_Berl = 73# * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.031 * flujo_másico_líquido - 0.063) * flujo_másico_líquido ^ -0.359
        End If
    Case 1.5
        If flujo_másico_líquido >= 0.68 And flujo_másico_líquido < 2 Then
        Área_interfacial_agua_Berl = 40.6 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (-0.058) * flujo_másico_líquido ^ 0.455
        ElseIf flujo_másico_líquido >= 2 And flujo_másico_líquido < 6.1 Then
        Área_interfacial_agua_Berl = 62.4 * (808 * flujo_másico_gas / densidad_gas ^ 0.5) ^ (0.024 * flujo_másico_líquido - 0.0996) * flujo_másico_líquido ^ -0.1355
        End If
End Select
End Function

'Cálculo de la retención estática para anillos Raschig de cerámica
Function Retención_estática_Raschig_cerámica(viscosidad_líquido, densidad_líquido, tensión_superficial_líquido, diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Retención_estática_Raschig_cerámica = 0.0486 * viscosidad_líquido ^ 0.02 * tensión_superficial_líquido ^ 0.99 / 0.01774 ^ 1.21 / densidad_líquido ^ 0.37
    Case 1
    Retención_estática_Raschig_cerámica = 0.0486 * viscosidad_líquido ^ 0.02 * tensión_superficial_líquido ^ 0.99 / 0.0356 ^ 1.21 / densidad_líquido ^ 0.37
    Case 1.5
    Retención_estática_Raschig_cerámica = 0.0486 * viscosidad_líquido ^ 0.02 * tensión_superficial_líquido ^ 0.99 / 0.053 ^ 1.21 / densidad_líquido ^ 0.37
    Case 2
    Retención_estática_Raschig_cerámica = 0.0486 * viscosidad_líquido ^ 0.02 * tensión_superficial_líquido ^ 0.99 / 0.0725 ^ 1.21 / densidad_líquido ^ 0.37
End Select
End Function

'Cálculo de la retención estática para anillos Raschig de carbón
Function Retención_estática_Raschig_carbón(viscosidad_líquido, densidad_líquido, tensión_superficial_líquido, diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 1
    Retención_estática_Raschig_carbón = 0.0237 * viscosidad_líquido ^ 0.02 * tensión_superficial_líquido ^ 0.23 / 0.01301 ^ 1.21 / densidad_líquido ^ 0.37
    Case 1.5
    Retención_estática_Raschig_carbón = 0.0237 * viscosidad_líquido ^ 0.02 * tensión_superficial_líquido ^ 0.23 / 0.0543 ^ 1.21 / densidad_líquido ^ 0.37
    Case 2
    Retención_estática_Raschig_carbón = 0.0237 * viscosidad_líquido ^ 0.02 * tensión_superficial_líquido ^ 0.23 / 0.0716 ^ 1.21 / densidad_líquido ^ 0.37
End Select
End Function

'Cálculo de la retención estática para sillas de Berl de cerámica
Function Retención_estática_Berl_cerámica(viscosidad_líquido, densidad_líquido, tensión_superficial_líquido, diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Retención_estática_Berl_cerámica = 0.00423 * viscosidad_líquido ^ 0.04 * tensión_superficial_líquido ^ 0.55 / 0.01622 ^ 1.56 / densidad_líquido ^ 0.37
    'parece tener un error
    Case 1
    Retención_estática_Berl_cerámica = 0.00423 * viscosidad_líquido ^ 0.04 * tensión_superficial_líquido ^ 0.55 / 0.032 ^ 1.56 / densidad_líquido ^ 0.37
    Case 1.5
    Retención_estática_Berl_cerámica = 0.00423 * viscosidad_líquido ^ 0.04 * tensión_superficial_líquido ^ 0.55 / 0.0472 ^ 1.56 / densidad_líquido ^ 0.37
End Select
End Function

'Cálculo de la retención estática para anillos Raschig de cerámica en agua
Function Retención_estática_Raschig_cerámica_agua(diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Retención_estática_Raschig_cerámica_agua = 0.000247 / 0.01774 ^ 1.21
    Case 1
    Retención_estática_Raschig_cerámica_agua = 0.000247 / 0.0356 ^ 1.21
    Case 1.5
    Retención_estática_Raschig_cerámica_agua = 0.000247 / 0.053 ^ 1.21
    Case 2
    Retención_estática_Raschig_cerámica_agua = 0.000247 / 0.0725 ^ 1.21
End Select
End Function

'Cálculo de la retención estática para anillos Raschig de carbón en agua
Function Retención_estática_Raschig_carbón_agua(diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 1
    Retención_estática_Raschig_carbón_agua = 0.000594 / 0.01301 ^ 1.21
    Case 1.5
    Retención_estática_Raschig_carbón_agua = 0.000594 / 0.0543 ^ 1.21
    Case 2
    Retención_estática_Raschig_carbón_agua = 0.000594 / 0.0716 ^ 1.21
End Select
End Function

'Cálculo de la retención estática para sillas de Berl de cerámica en agua
Function Retención_estática_Berl_cerámica_agua(diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Retención_estática_Berl_cerámica_agua = 0.00005014 / 0.31622 ^ 1.56
    'parece tener un error
    Case 1
    Retención_estática_Berl_cerámica_agua = 0.00005014 / 0.032 ^ 1.56
    Case 1.5
    Retención_estática_Berl_cerámica_agua = 0.00005014 / 0.0472 ^ 1.56
End Select
End Function

'Cálculo de la retención total para anillos Raschig de cerámica en agua
Function Retención_total_Raschig_cerámica_agua(flujo_másico_líquido, diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Retención_total_Raschig_cerámica_agua = 0.00000209 * (737.5 * flujo_másico_líquido) ^ (1.508 * 0.01774 ^ 0.376) / 0.01774 ^ 2
    Case 1
    Retención_total_Raschig_cerámica_agua = 0.00000209 * (737.5 * flujo_másico_líquido) ^ (1.508 * 0.0356 ^ 0.376) / 0.0356 ^ 2
    Case 1.5
    Retención_total_Raschig_cerámica_agua = 0.00000209 * (737.5 * flujo_másico_líquido) ^ (1.508 * 0.053 ^ 0.376) / 0.053 ^ 2
    Case 2
    Retención_total_Raschig_cerámica_agua = 0.00000209 * (737.5 * flujo_másico_líquido) ^ (1.508 * 0.0725 ^ 0.376) / 0.0725 ^ 2
End Select
End Function

'Cálculo de la retención total para anillos Raschig de carbón en agua
Function Retención_total_Raschig_carbón_agua(flujo_másico_líquido, diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 1
    Retención_total_Raschig_carbón_agua = 0.00000734 * (737.5 * flujo_másico_líquido) ^ (1.104 * 0.01301 ^ 0.376) / 0.013016 ^ 2
    Case 1.5
    Retención_total_Raschig_carbón_agua = 0.00000734 * (737.5 * flujo_másico_líquido) ^ (1.104 * 0.0543 ^ 0.376) / 0.0543 ^ 2
    Case 2
    Retención_total_Raschig_carbón_agua = 0.00000734 * (737.5 * flujo_másico_líquido) ^ (1.104 * 0.0716 ^ 0.376) / 0.0716 ^ 2
End Select
End Function
   
'Cálculo de la retención total para sillas de Berl de cerámica en agua
Function Retención_total_Berl_cerámica_agua(flujo_másico_líquido, diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Retención_total_Berl_cerámica_agua = 0.00000232 * (737.5 * flujo_másico_líquido) ^ (1.508 * 0.31622 ^ 0.376) / 0.31622 ^ 2
    'parece tener un error
    Case 1
    Retención_total_Berl_cerámica_agua = 0.00000232 * (737.5 * flujo_másico_líquido) ^ (1.508 * 0.032 ^ 0.376) / 0.032 ^ 2
    Case 1.5
    Retención_total_Berl_cerámica_agua = 0.00000232 * (737.5 * flujo_másico_líquido) ^ (1.508 * 0.0472 ^ 0.376) / 0.0472 ^ 2
End Select
End Function

'Cálculo del área interfacial
Function Área_interfacial(Área_interfacial_agua, retención_operativa, retención_operativa_agua)
Área_interfacial = Área_interfacial_agua * retención_operativa / retención_operativa_agua
End Function

'Cálculo del espacio vacío seco de un lecho de anillos Raschig de cerámica
'Colocar el diámetro nominal en pulgadas
Function Espacio_vacío_seco_Raschig_cerámica(diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Espacio_vacío_seco_Raschig_cerámica = 0.63
    Case 1
    Espacio_vacío_seco_Raschig_cerámica = 0.73
    Case 1.5
    Espacio_vacío_seco_Raschig_cerámica = 0.71
    Case 2
    Espacio_vacío_seco_Raschig_cerámica = 0.74
End Select
End Function

'Cálculo del espacio vacío seco de un lecho de sillas de Berl de cerámica
'Colocar el diámetro nominal en pulgadas
Function Espacio_vacío_seco_Berl_cerámica(diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Espacio_vacío_seco_Berl_cerámica = 0.63
    Case 1
    Espacio_vacío_seco_Berl_cerámica = 0.69
    Case 1.5
    Espacio_vacío_seco_Berl_cerámica = 0.75
    Case 2
    Espacio_vacío_seco_Berl_cerámica = 0.72
End Select
End Function

'Cálculo del factor de empaque para anillos Raschig
Function Factor_empaque_Raschig(diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Factor_empaque_Raschig = 580
    Case 1
    Factor_empaque_Raschig = 155
    Case 1.5
    Factor_empaque_Raschig = 95
    Case 2
    Factor_empaque_Raschig = 65
End Select
End Function

'Cálculo del factor de empaque para sillas de Berl
Function Factor_empaque_Berl(diámetro_relleno_in)
Select Case diámetro_relleno_in
    Case 0.5
    Factor_empaque_Berl = 240
    Case 1
    Factor_empaque_Berl = 110
    Case 1.5
    Factor_empaque_Berl = 65
    Case 2
    Factor_empaque_Berl = 45
End Select
End Function

'Cálculo del número de etapas teóricas para una torre de absorción de platos en un sistema diluido (Ecuación de Colburn)
Function N_abs(factor_absorción, constante_Henry, x_0, y_1, y_Nplus1)
If factor_absorción = 1 Then
N_abs = (y_Nplus1 - y_1) / (y_1 - constante_Henry * x_0)
Else
N_abs = Log(1 / factor_absorción + (1 - 1 / factor_absorción) * (y_Nplus1 - constante_Henry * x_0) / (y_1 - constante_Henry * x_0)) / Log(factor_absorción)
End If
End Function

'Cálculo del número de etapas teóricas para una torre de desorción de platos en un sistema diluido (Ecuación de Colburn)
Function N_des(factor_desorción, constante_Henry, x_0, x_N, y_S0)
If factor_desorción = 1 Then
N_des = (x_N - x_0) / (x_0 - y_S0 / constante_Henry)
Else
N_des = Log(1 / factor_desorción + (1 - 1 / factor_desorción) * (x_N - y_S0 / constante_Henry) / (x_0 - y_S0 / constante_Henry)) / Log(factor_desorción)
End If
End Function

'Cálculo del número de unidades globales de transferencia referido a la fase gaseosa para torres de absorción rellenas, sistema diluido
Function NOG_abs(factor_absorción, constante_Henry, x_0, y_1, y_Nplus1)
If factor_absorción = 1 Then
NOG_abs = (y_Nplus1 - y_1) / (y_1 - constante_Henry * x_0)
Else
NOG_abs = Log(1 / factor_absorción + (1 - 1 / factor_absorción) * (y_Nplus1 - constante_Henry * x_0) / (y_1 - constante_Henry * x_0)) / (1 - 1 / factor_absorción)
End If
End Function

'Cálculo del número de unidades globales de transferencia referido a la fase líquida para torres de desorción rellenas, sistema diluido
Function NOL_des(factor_desorción, constante_Henry, x_N, x0, y_S0)
If factor_desorción = 1 Then
NOL_des = (x_N - x_0) / (x_0 - y_S0 / constante_Henry)
Else
NOL_des = Log(1 / factor_desorción + (1 - 1 / factor_desorción) * (x_N - y_S0 / constante_Henry) / (x_0 - y_S0 / constante_Henry)) / (1 - 1 / factor_desorción)
End If
End Function

'Cálculo de la ordenada de la correlación de Eckert para la condición de inundación
Function Ord_Eckert_inundación(Abs_Eckert)
Ord_Eckert_inundación = Exp(-3.428 - 1.062 * Log(Abs_Eckert) - 0.121 * Log(Abs_Eckert) ^ 2)
End Function

'Cálculo de la abscisa de la correlación de Eckert
Function Abs_Eckert(caudal_másico_líquido, caudal_másico_gas, densidad_líquido, densidad_gas)
Abs_Eckert = caudal_másico_líquido / caudal_másico_gas * (densidad_gas / (densidad_líquido - densidad_gas)) ^ 0.5
End Function

'Cálculo de la ordenada de la correlación de Eckert para una caída de presión de
'200 Pa/m
Function Ord_Eckert_200(Abs_Eckert)
Ord_Eckert_200 = Exp(-4.472 - 0.865 * Log(Abs_Eckert) - 0.127 * Log(Abs_Eckert) ^ 2)
End Function

'Cálculo de la ordenada de la correlación de Eckert para una caída de presión de
'400 Pa/m
Function Ord_Eckert_400(Abs_Eckert)
Ord_Eckert_400 = Exp(-4.0781 - 0.8908 * Log(Abs_Eckert) - 0.1229 * Log(Abs_Eckert) ^ 2)
End Function

'Cálculo de la difusividad en fase gaseosa del sistema amoníaco-aire
Function Difusividad_amoníaco_aire(temperatura_abs, presión)
Difusividad_amoníaco_aire = 0.001 * temperatura_abs ^ 1.75 * (1 / 17 + 1 / 28.8) ^ 0.5 / presión / (20.1 ^ (1 / 3) + 14.9 ^ (1 / 3)) ^ 2 / 10000
End Function

'Cálculo de la difusividad en fase líquida del sistema amoníaco-agua
Function Difusividad_amoníaco_agua(temperatura_abs)
'Ec. 2-159 Perry 7ª Ed., la correlación es válida para sistemasde hasta 10% molar de
'soluto en agua
'Puede fallar!
Difusividad_amoníaco_agua = 8.621E-14 / Exp(-5.2886 + 1564.8 / (temperatura_abs)) / (17.03 / (111.5 / 0.082 / 405.5 / 0.2466 ^ (1 + (1 - 239.8 / 405.5) ^ (2 / 7)))) ^ 0.589
End Function

'Cálculo de la viscosidad del agua
Function Viscosidad_agua(temperatura_abs)
Viscosidad_agua = Exp(-5.2886 + 1564.8 / (temperatura_abs))
End Function

'Cálculo de la fracción molar de equilibrio del amoníaco en fase gaseosa para el sistema
'amoníaco-agua a 4.4ºC en función de la concentración del mismo en fase líquida
Function y_eq_amoníaco_agua_5degC(x)
y_eq_amoníaco_agua_5degC = 17.148 * x ^ 3 - 3.2958 * x ^ 2 + 0.6195 * x
End Function

'Cálculo de la densidad de la solución amoniacal para 5ºC
Function Densidad_sol_amoniacal_5degC(x)
Densidad_sol_amoniacal_5degC = -319.16 * x + 998.54
End Function

'Cálculo de la curva de equilibrio para un sistema binario con volatilidad relativa
'constante
Function y_eq_vapor(x_eq_líquido, volatilidad_relativa)
If x_eq_líquido = 0 Then
y_eq_vapor = 0
Else
y_eq_vapor = 1 / (1 + (1 / x_eq_líquido - 1) / volatilidad_relativa)
End If
End Function

'Cálculo de la línea de operaciones de la zona de agotamiento según método
'McCabe-Thiele
Function y_loza(x, x_destilado, R)
y_loza = R / (R + 1) * x + x_destilado / (R + 1)
End Function

'Cálculo de la línea de alimentación según método McCabe-Thiele
Function y_alimentación(x, x_alimentación, q)
y_alimentación = q / (q - 1) * x - x_alimentación / (q - 1)
End Function

'Cálculo del número mínimo de platos para un sistema  binario de volatilidad relativa constante (Ecuación de Fenske).
Function N_mínimo_Fenske(x_destilado, x_residuo, volatilidad_relativa_promedio)
N_mínimo_Fenske = Log(x_destilado * (1 - x_residuo) / x_residuo / (1 - x_destilado)) / Log(volatilidad_relativa_promedio)
End Function

'Cálculo del número mínimo de platos para un sistema multicomponente de volatilidad relativa constante (Ecuación de Fenske aplicada a multicomponentes).
Function N_mínimo_Fenske_multicomponente(x_claveliviano_destilado, x_clavepesado_destilado, x_claveliviano_residuo, x_clavepesado_residuo, volatilidad_relativa_claveliviano_promedio)
N_mínimo_Fenske_multicomponente = Log(x_claveliviano_destilado * x_clavepesado_residuo / x_claveliviano_residuo / x_clavepesado_destilado) / Log(volatilidad_relativa_claveliviano_promedio)
End Function


'Cálculo del número de platos teóricos necesarios en función del número mínimo de platos, y las relaciones de reflujo mínima y operativa (Correlación de Gilliland).
Function N_Gilliland(R_mínimo, R_operativo, N_mínimo)
x = (R_operativo - R_mínimo) / (R_operativo + 1)
y = 1 - Exp((1 + 54.4 * x) * (x - 1) / (11 + 117.2 * x) / x ^ 0.5)
N_Gilliland = (y + N_mínimo) / (1 - y)
End Function

'Determinación analítica de la relación mínima de reflujo para un sistema binario
'de volatilidad relativa constante (Fórmula de Underwood)
Function Rmín_Underwood(x_alimentación, x_destilado, q, volatilidad_relativa, semilla, tolerancia)
R_mínimo = semilla
diferencia_abs = tolerancia * 10
MI = (R_mínimo * x_alimentación + q * x_destilado) / (R_mínimo * (1 - x_alimentación) + q * (1 - x_destilado))
MD = volatilidad_relativa * (x_destilado * (q - 1) + x_alimentación * (R_mínimo + 1)) / ((R_mínimo + 1) * (1 - x_alimentación) + (q - 1) * (1 - x_destilado))
Do While diferencia_abs > tolerancia
    y = (R_mínimo * x_alimentación + q * x_destilado) / (R_mínimo * (1 - x_alimentación) + q * (1 - x_destilado)) - volatilidad_relativa * (x_destilado * (q - 1) + x_alimentación * (R_mínimo + 1)) / ((R_mínimo + 1) * (1 - x_alimentación) + (q - 1) * (1 - x_destilado))
    y_dy = ((R_mínimo + 0.0001) * x_alimentación + q * x_destilado) / ((R_mínimo + 0.0001) * (1 - x_alimentación) + q * (1 - x_destilado)) - volatilidad_relativa * (x_destilado * (q - 1) + x_alimentación * (R_mínimo + 0.0001 + 1)) / ((R_mínimo + 0.0001 + 1) * (1 - x_alimentación) + (q - 1) * (1 - x_destilado))
    derivada = (y_dy - y) / 0.0001
    R_mínimo = R_mínimo - y / derivada
    diferencia = y / derivada
    diferencia_abs = Abs(diferencia)
Loop
Rmín_Underwood = R_mínimo
End Function


'Determinación analítica de la relación mínima de reflujo para un sistema
'multicomponente (Fórmula de Underwood modificada para multicomponentes)
Function R_mínimo_Underwood_multicomponentes(x1_alimentación, x2_alimentación, x3_alimentación, x4_alimentación, x1_destilado, x2_destilado, x3_destilado, x4_destilado, volatilidad_relativa_1, volatilidad_relativa_2, volatilidad_relativa_3, volatilidad_relativa_4, q, semilla, tolerancia)
theta = semilla
diferencia_abs = tolerancia * 10
Do While diferencia_abs > tolerancia
    y = volatilidad_relativa_1 * x1_alimentación / (volatilidad_relativa_1 - theta) + volatilidad_relativa_2 * x2_alimentación / (volatilidad_relativa_2 - theta) + volatilidad_relativa_3 * x3_alimentación / (volatilidad_relativa_3 - theta) + volatilidad_relativa_4 * x4_alimentación / (volatilidad_relativa_4 - theta) + q - 1
    y_dy = volatilidad_relativa_1 * x1_alimentación / (volatilidad_relativa_1 - theta - 0.0001) + volatilidad_relativa_2 * x2_alimentación / (volatilidad_relativa_2 - theta - 0.0001) + volatilidad_relativa_3 * x3_alimentación / (volatilidad_relativa_3 - theta - 0.0001) + volatilidad_relativa_4 * x4_alimentación / (volatilidad_relativa_4 - theta - 0.0001) + q - 1
    derivada = (y_dy - y) / 0.0001
    theta = theta - y / derivada
    diferencia = y / derivada
    diferencia_abs = Abs(diferencia)
Loop
R_mínimo_Underwood_multicomponentes = volatilidad_relativa_1 * x1_destilado / (volatilidad_relativa_1 - theta) + volatilidad_relativa_2 * x2_destilado / (volatilidad_relativa_2 - theta) + volatilidad_relativa_3 * x3_destilado / (volatilidad_relativa_3 - theta) + volatilidad_relativa_4 * x4_destilado / (volatilidad_relativa_4 - theta) - 1
End Function

'Cálculo de la ubicación del plato de alimentación según Kirkbridge
Function Nalim(N, W, D, x_claveliviano_alimentación, x_clavepesado_alimentación, x_claveliviano_residuo, x_clavepesado_destilado)
Nalim = N / (1 + (D / W * (x_claveliviano_alimentación / x_clavepesado_alimentación) * (x_clavepesado_destilado / x_claveliviano_residuo) ^ 2) ^ -0.206)
End Function



'Cálculo de la velocidad de inundación referida al área neta
'Aplica para diámetros de orificio <0.006m
Function Velocidad_inundación_ref_área_neta(densidad_líquido, densidad_gas, tensión_superficial, espaciamiento_platos, relación_másica_L_G, Frac_Aactiva_ocup_orificios, diámetro_orificio)
J = relación_másica_L_G * (densidad_gas / densidad_líquido) ^ 0.5
alfa = 0.0744 * espaciamiento_platos + 0.01173
beta = 0.0304 * espaciamiento_platos + 0.015
If diámetro_orificio >= 0.006 Then
Velocidad_inundación_ref_área_neta = "*diám.orificio"
Else
If Frac_Aactiva_ocup_orificios > 0.1 Then
If J < 1 Then
Velocidad_inundación_ref_área_neta = (beta + alfa) * (tensión_superficial / 0.02) ^ 0.2 * ((densidad_líquido - densidad_gas) / densidad_gas) ^ 0.5
Else
Velocidad_inundación_ref_área_neta = (beta - alfa * 2.3 * Log(J)) * (tensión_superficial / 0.02) ^ 0.2 * ((densidad_líquido - densidad_gas) / densidad_gas) ^ 0.5
End If
Else
Velocidad_inundación_ref_área_neta = (5 * Frac_Aactiva_ocup_orificios + 0.5) * (beta - alfa * 2.3 * Log(J)) * (tensión_superficial / 0.02) ^ 0.2 * ((densidad_líquido - densidad_gas) / densidad_gas) ^ 0.5
End If
End If
End Function

'Cálculo de la fracción de área total ocupada por un vertedero
Function Fracción_Atotal_ocup_vertedero(relación_W_T)
Select Case relación_W_T
    Case 0.55
    Fracción_Atotal_ocup_vertedero = 0.03877
    Case 0.6
    Fracción_Atotal_ocup_vertedero = 0.05257
    Case 0.65
    Fracción_Atotal_ocup_vertedero = 0.06899
    Case 0.7
    Fracción_Atotal_ocup_vertedero = 0.08808
    Case 0.75
    Fracción_Atotal_ocup_vertedero = 0.11255
    Case 0.8
    Fracción_Atotal_ocup_vertedero = 0.14145
    Case Else
    Fracción_Atotal_ocup_vertedero = "W/T no disp"
End Select
End Function

'Cálculo del espaciamiento entre platos recomendado
Function Espaciamiento_platos_recomendado(diámetro_torre)
Select Case diámetro_torre
    Case Is < 1
    Espaciamiento_platos_recomendado = 0.5
    Case Is >= 1 < 3
    Espaciamiento_platos_recomendado = 0.6
    Case Is >= 3 < 4
    Espaciamiento_platos_recomendado = 0.75
    Case Is >= 4 < 8
    Espaciamiento_platos_recomendado = 0.9
End Select
End Function

'Cálculo de la fracción de área activa ocupada por orificios
Function Frac_Aactiva_ocup_orificios(diámetro_orificio, paso)
Frac_Aactiva_ocup_orificios = 0.907 * (diámetro_orificio / paso) ^ 2
End Function

'Cálculo de la altura de la ola de líquido sobre el vertedero
Function Altura_ola_líquido_sobre_vertedero(caudal_volumétrico_líquido, longitud_vertedero)
Altura_ola_líquido_sobre_vertedero = 0.666 * (caudal_volumétrico_líquido / longitud_vertedero) ^ (2 / 3)
End Function

'Cálculo de la eficiencia puntual del plato referido a la fase gaseosa
Function Eficiencia_puntual_plato_ref_gas(constante_Henry, relación_molar_L_G, altura_vertedero, velocidad_área_activa, densidad_gas, caudal_volumétrico_líq, distancia_entre_vertederos, Schmidt_gas, difusividad_líq, altura_ola_líq, diámetro_torre, longitud_vertedero)
NtG = (0.776 + 4.57 * altura_vertedero - 0.238 * velocidad_área_activa * densidad_gas ^ 0.5 + 104.6 * caudal_volumétrico_líq / distancia_entre_vertederos) / Schmidt_gas ^ 0.5
tetha = (altura_vertedero + altura_ola_líq) * (diámetro_torre + longitud_vertedero) * distancia_entre_vertederos * 0.5 / caudal_volumétrico_líq
NtL = 40000 * difusividad_líq ^ 0.5 * tetha * (0.213 * velocidad_área_activa * densidad_gas ^ 0.5 + 0.15)
NtOG = 1 / (1 / NtG + constante_Henry / NtL / relación_molar_L_G)
Eficiencia_puntual_plato_ref_gas = 1 - Exp(-NtOG)
'Funciona correctamente, comprobado con ejemplo resuelto de Treybal
End Function

'Function tetha(altura_vertedero, altura_ola_líq, diámetro_torre, longitud_vertedero, distancia_entre_vertederos, caudal_volumétrico_líq)
'tetha = (altura_vertedero + altura_ola_líq) * (diámetro_torre + longitud_vertedero) * distancia_entre_vertederos * 0.5 / caudal_volumétrico_líq
'End Function
'Funciona correctamente, comprobado con ejemplo resuelto de treibal
'En el texto se utiliza hL en lugar de h1+hw

'Function NtG(altura_vertedero, velocidad_área_activa, densidad_gas, caudal_volumétrico_líq, distancia_entre_vertederos, Schmidt_gas)
'NtG = (0.776 + 4.57 * altura_vertedero - 0.238 * velocidad_área_activa * densidad_líq ^ 0.5 + 104.6 * caudal_volumétrico_líq / distancia_entre_vertederos) / Schmidt_gas ^ 0.5
'End Function
'Funciona correctamente, comprobado con ejemplo resuelto de treibal

'Function NtL(difusividad_líq, tetha, velocidad_área_activa, densidad_gas)
'NtL = 40000 * difusividad_líq ^ 0.5 * tetha * (0.213 * velocidad_área_activa * densidad_gas ^ 0.5 + 0.15)
'End Function
'Funciona correctamente, comprobado con ejemplo resuelto de treibal


'Function DE(velocidad_área_activa, caudal_volumétrico_líq, distancia_entre_vertederos, altura_vertedero)
'DE = (0.00393 + 0.0171 * velocidad_área_activa + 3.67 * caudal_volumétrico_líq / distancia_entre_vertederos + 0.18 * altura_vertedero) ^ 2
'End Function

'Function Pe(distancia_entre_vertederos, DE, tetha)
'Pe = distancia_entre_vertederos ^ 2 / DE / tetha
'End Function
'Funciona correctamente

'Function etha(Pe, constante_Henry, Eficiencia_puntual_plato_ref_gas, relación_molar_L_G)
'etha = 0.5 * Pe * ((1 + 4 * constante_Henry * Eficiencia_puntual_plato_ref_gas / relación_molar_L_G / Pe) ^ 0.5 - 1)
'End Function
'Funciona correctamente

'Cálculo de la eficiencia de plato de Murphree
Function Eficiencia_plato_Murphree(Eficiencia_puntual_plato_ref_gas, constante_Henry, relación_molar_L_G, distancia_entre_vertederos, altura_vertedero, altura_ola_líq, diámetro_torre, longitud_vertedero, caudal_volumétrico_líq, velocidad_área_activa)
DE = (0.00393 + 0.0171 * velocidad_área_activa + 3.67 * caudal_volumétrico_líq / distancia_entre_vertederos + 0.18 * altura_vertedero) ^ 2
tetha = (altura_vertedero + altura_ola_líq) * (diámetro_torre + longitud_vertedero) * distancia_entre_vertederos * 0.5 / caudal_volumétrico_líq
Pe = distancia_entre_vertederos ^ 2 / DE / tetha
etha = 0.5 * Pe * ((1 + 4 * constante_Henry * Eficiencia_puntual_plato_ref_gas / relación_molar_L_G / Pe) ^ 0.5 - 1)
Eficiencia_plato_Murphree = Eficiencia_puntual_plato_ref_gas * ((1 - Exp(-etha - Pe)) / (etha + Pe) / (1 + (etha + Pe) / etha) + (Exp(etha) - 1) / etha / (1 + etha / (etha + Pe)))
End Function
'Funciona correctamente

'Cálculo de la eficiencia de plato de Murphree corregida por arrastre
Function Eficiencia_plato_Murphree_corregida_arrastre(Eficiencia_plato_Murphree, arrastre_líquido)
Eficiencia_plato_Murphree_corregida_arrastre = Eficiencia_plato_Murphree / (1 + Eficiencia_plato_Murphree * (arrastre_líquido / (1 - arrastre_líquido)))
End Function

'Cálculo del arrastre fraccionario
Function Arrastre_fraccionario(relación_másica_L_G, densidad_gas, densidad_líq, velocidad_área_neta_ref_inundación)
AbsCorr = relación_másica_L_G * (densidad_gas / densidad_líq) ^ 0.5
constante = Exp(-0.1868 * Log(AbsCorr) ^ 2 - 2.2944 * Log(AbsCorr) - 7.1214)
exponente = Exp(-0.0785 * Log(AbsCorr) ^ 2 - 0.5793 * Log(AbsCorr) + 0.2809)
Arrastre_fraccionario = constante * velocidad_área_neta_ref_inundación ^ exponente
End Function
'Correlación válida para V/Vf en el rango 0.6-0.9

'Función J dependiente de alfa y beta
Function mibesselI(x)
mibesselI = Application.WorksheetFunction.BesselI(x, 0)
End Function
Function J(alfa, beta)
zN = alfa
nn = 1000
deltaz = zN / nn
For i = 0 To nn - 1
Z = i * deltaz
z1 = Evaluate(mibesselI(2 * Sqr(beta * Z))) * Exp(-Z)
z2 = Evaluate(mibesselI(2 * Sqr(beta * (Z + deltaz)))) * Exp(-(Z + deltaz))
integral = integral + deltaz * (z1 + z2) / 2
Next
J = 1 - integral * Exp(-beta)
End Function


'Cálculo de la integral definida por método numérico de una función dependiente de la variable "x"
Function integral_definida(fórmula_string As String, x_string As String, x_inf, x_sup, N)
Dim x As Double
Dim i As Integer
Dim fórmula_string_int As String
dx = (x_sup - x_inf) / N
   For i = 0 To N - 1
        x = x_inf + i * dx
        x1 = Replace(CStr(x), ",", ".")
        x2 = Replace(CStr(x + dx), ",", ".")
    fórmula_string_int = Replace(fórmula_string, ",", ".")
        y1 = Evaluate(Replace(fórmula_string_int, "x", x1))
        y2 = Evaluate(Replace(fórmula_string_int, "x", x2))
        integral_definida = integral_definida + dx * (y1 + y2) / 2
    Next
End Function

'Cálculo del área bajo una curva ajustando con el trapecio
Function área(valoresX, valoresY)
N = valoresX.Count
For J = 2 To N
área = área + (valoresX(J) - valoresX(J - 1)) * (valoresY(J) + valoresY(J - 1)) / 2
Next J
End Function

'Calcula la derivada a partir de la pendiente de la recta secante en el caso límite
Function derivada(fórmula_string As String, x_string As String, x_0, dx)
Dim fórmula_string_int As String
    fórmula_string_int = Replace(fórmula_string, ",", ".")
    y1 = Evaluate(Replace(fórmula_string_int, "x", x_0))
    x2 = x_0 + dx
    y2 = Evaluate(Replace(fórmula_string_int, "x", Replace(CStr(x2), ",", ".")))
    derivada = (y2 - y1) / dx
End Function

'Cálculo de raíces de ecuaciones por el método numérico de Newton
Function Hallar_raíz(fórmula_string As String, x_string As String, x_semilla As Double, tolerancia As Double)
Dim fórmula_string_int As String
Dim x_var, x_dx, y_dy, diferencia, derivada, diferencia_abs As Double
x_var = x_semilla
diferencia_abs = tolerancia * 10
fórmula_string_int = Replace(fórmula_string, ",", ".")
Do While diferencia_abs > tolerancia
    y = Evaluate(Replace(fórmula_string_int, "x", Replace(x_var, ",", ".")))
    x_dx = x_var + 0.0001
    y_dy = Evaluate(Replace(fórmula_string_int, "x", Replace(x_dx, ",", ".")))
    derivada = (y_dy - y) / 0.0001
    x_var = x_var - y / derivada
    diferencia = y / derivada
    diferencia_abs = Abs(diferencia)
Loop
Hallar_raíz = x_var
End Function

'Muestra la fórmula ingresada en la celda seleccionada
Function MostrarFórmula(celda)
MostrarFórmula = celda.Formula
End Function


