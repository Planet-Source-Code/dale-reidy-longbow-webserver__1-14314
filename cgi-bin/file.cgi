redefstr data
if (method$ == 'write')
 file openout 'test.dat' 1
 fwrite 1 data$
 file close 1
 print '<META HTTP-EQUIV="Refresh" Content="1; URL=update.html">' 
else
 file openin 'test.dat' 1
 fread 1 data$
 file close 1
 print 'document.write("' & data$ & '")'
endif