if (name$ == 'dale') || (name$ == 'bob') && (password$ == 'me')
	print '<html>'
	print '<body bgcolor=red>'
	print '<font face=verdana color=orange>'
	print '<b>Welcome ' & name$ & ' to this test page'
	goto CARRY_THIS_ON
	print '</font></body></html>'
else
//if (name$ != 'dale') && (name$ != 'bob') && (password$ != 'me')

	print 'ACCESS DENIED'
endif

end

CARRY_THIS_ON>
	print '<b>GOTO STATUS:OK</b>'
end