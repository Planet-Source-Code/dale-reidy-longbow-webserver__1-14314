newstr data
file openin 'test2.dat' 1
again>
fread 1 data$
print data$
if (eof(1) == 0)
	goto again
endif
file close 1