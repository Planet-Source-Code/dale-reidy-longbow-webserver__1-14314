newint incount

file openin 'counter.dat' 1
fread 1 incount%
file close 1

inc incount%

file openout 'counter.dat' 2
fwrite 2 incount%
file close 2

print 'document.write("' & incount% & '")'