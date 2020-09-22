newstr newname
newstr newpass
newstr newpass2
newstr temp1
newstr temp2

// Check To Make Sure All The Information Was Entered
if (newname$ == '') || (newpass$ == '') || (newpass2$ == '')
	print 'All Data Not Entered'
	end
endif


// Check To Make Sure Both Passwords Match
if (newpass$ != newpass2$)
	print 'Passwords Do Not Match'
	end
endif

// Check To Make Sure The User Doesn't Already Exist
file openin 'login.dtx' 1
again>
fmread 1 temp1$ temp2$
if (temp1$ == newname$)
	print 'User Already Exists'
	file close 1
	end
endif
if (eof(1) == 0)
	goto again
endif

file close 1

file openappend 'login.dtx' 1
temp1$ = newname$ & ',' & newpass$ & lbCrLf
fwrite 1 temp1$
file close 1

print '<html><title>New User:' & newname$ & ' Created.</title>'
print '<body>'
print '<b>New User ' & newname$ & ' Created.</b><br>'
print '<br><br>'
print '<a href=index.html>Click Here To Return To The Login Page</a>'
print '</body></html>'