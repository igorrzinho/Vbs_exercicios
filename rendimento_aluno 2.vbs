Dim n1,n2,media, resp, situacao 'declaração de variaveis'
Dim audio 
call carregar_voz'chama uma função'
sub carregar_voz()
	set audio = CreateObject("SAPI.SPVOICE")
	audio.volume = 100
	audio.rate = 2 'velocidade da voz'
	call entrada_notas 
end Sub


sub entrada_notas()
	'entrada de dados'
	n1=cdbl(inputbox("digite a nota 1:","rendimentos aluno"))
	n2=cdbl(inputbox("digite a nota 2:","rendimentos aluno"))
	'processamento'
	media = round((n1 + n2) / 2, 1)
	if media < 4 then
		situacao = "reprovado"
	elseif media >= 4 and media < 6 then
		situacao = "recuperação"
	else 
		situacao = "aprovado"
	end if
	'saida por voz'
	audio.speak("Média do aluno"& media &""+vbNewLine&_
				"situação do aluno"& situacao &"")
	'saida de dados por mensagem'
	resp = msgbox("media do aluno: "& media &"" + vbNewLine&_
				 "situacao do aluno: "& situacao &"" + vbNewLine&_
				 "novo calculo?", vbQuestion+vbYesNo, "Rend. Aluno")
	if resp = vbYes then ''
		call entrada_notas
	else 
		MsgBox("fim do programa")
	end if
	
end Sub	
		