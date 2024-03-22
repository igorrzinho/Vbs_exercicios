Dim qtdsalario, salbruto, salliq, impo, resp

call sal
sub sal()
	qtdsalario = cdbl(InputBox("digite o numero de salarios minimos recebidos(1412):"))
	salbruto = qtdsalario * 1412
	if qtdsalario >= 2 Then
		salliq = (salbruto / 100 ) * 89
	else 
		salliq = salbruto
	end if
	impo = salbruto - salliq

	resp = MsgBox("salario bruto "& salbruto &"" + vbNewLine&_
				  "salario liquido "& salliq &"" + vbNewLine&_
				  "impostos "& impo &"" )
end Sub