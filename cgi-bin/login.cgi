newstr una1
newstr ups1

file openin 'login.dtx' 1
again>
fmread 1 una1$ ups1$

if (una1$ == username$) && (ups1$ == password$)
	includefile '23423480923423423049823.txt'
	end
endif

if (eof(1) == 0)
	goto again
endif

fclose 1

print 'ACCESS DENIED'