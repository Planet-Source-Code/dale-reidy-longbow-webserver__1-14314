
// Make sure the user entered a password
if (password$ != 'Howdy')
	if (nom$ != '')
		print '<html><title>Access Denied</title>'
		print '<body bgcolor=black>'
		print '<font face=verdana color=orange>'
		print '<b>Access Denied</b><br>'
		print '</body></html>'
		end
	endif
endif

// Show screen depending on value of nom
if (nom$ != '')
	print '<html><title>Welcome ' & nom$ & '</title>'
	print '<body bgcolor=black>'
	print '<font face=verdana color=orange>'
	print 'Welcome ' & nom$ & ' to dales test site'
	print '</font></body></html>'
endif

if (nom$ == '')
	print '<html><title>User Login</title>'
	print '<body bgcolor=black>'
	print '<font face=verdana color=orange>'
	print '<img src=islogo.jpg>'
	print '<br>'
	print '<b>......User Login......</b><br>'
	print '</font><font face=verdana color=orange size=-1>'
	print '<form method=post action=welcome.cgi>'
	print '<b>Username: </b>'
	print '<input type=text name=nom maxlength=30><br>'
	print '<b>Password: </b>'
	print '<input type=password name=password maxlength=10><br>'
	print '<input type=submit value=Login>'
	print '</form></font></body></html>'
endif