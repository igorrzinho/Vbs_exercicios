Dim numero, anumero, snumero, resp


call num
sub num
	numero = cint(InputBox("digite um numero:"))
	anumero = numero - 1
	snumero = numero + 1
	resp = MsgBox("antecessor "& anumero &"" + vbNewLine&_
				  "sucessor "& snumero &"" )
end Sub