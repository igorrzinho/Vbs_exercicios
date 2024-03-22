Dim lado, area, perimetro, resp

call calc_quadrado
sub calc_quadrado
	lado = cdbl(InputBox("digite tamanho do lado do quadrado:"))
	area = lado * lado
	perimetro = lado + lado + lado + lado
	resp = MsgBox("perimetro do quadrado "& perimetro &"" + vbNewLine&_
				  "area do quadrado "& area &"" )
end Sub