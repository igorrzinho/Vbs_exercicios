Dim n1,n2,n3,n4,apoio1, apoio2, manum, menum, resp


call num
sub num()
	n1 = cint(InputBox("digite um numero:"))
	n2 = cint(InputBox("digite um numero:"))
	n3 = cint(InputBox("digite um numero:"))
	n4 = cint(InputBox("digite um numero:"))
	if n1 > n2 Then
		apoio1 = n1
	else
		apoio1 = n2
	end if

	if n3 > n4 Then
		apoio2 = n3
		
	else
		apoio2 = n4
	end if

	if apoio1 > apoio2 Then
		manum = apoio1
	else
		manum = apoio2
	end if

	if n1 < n2 Then
		apoio1 = n1
	else 
		apoio1 = n2
	end if

	if n3 < n4 Then
		apoio2 = n3
	else 
		apoio2 = n4
	end if

	if apoio1 < apoio2 Then
		menum = apoio1
	Else
		menum = apoio2
	end if

	resp = MsgBox("maior "& manum &"" + vbNewLine&_
				  "menor "& menum &"" )
end Sub