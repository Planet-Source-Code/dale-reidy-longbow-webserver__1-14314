newstr first
newstr second

file openin 'name.dat' 1
fmread 1 first$ second$
file close 1

print first$ & '=' & second$