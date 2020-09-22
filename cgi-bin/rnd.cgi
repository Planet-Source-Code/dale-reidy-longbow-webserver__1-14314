newint random_number
newstr rndstr

rndstr$ = 10

strtoint rndstr$ random_number%

random_number% = random_number% - (1 * (3 * 3))

print random_number%