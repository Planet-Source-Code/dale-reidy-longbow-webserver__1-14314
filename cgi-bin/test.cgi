proc html_header
 print '<html><title>' & title$ & '</title><body>'
endproc

proc html_footer
 print '</body></html>'
endproc

proc access_denied
 title$ = 'ACCESS DENIED'
 call html_header
 print '<B>ACCESS DENIED</B>'
 call html_footer
endproc

proc access_granted
 call html_header
print '<font face=trebuchet color=orange>ACCESS GRANTED</font>'
 call html_footer
endproc

proc main
newstr title 'HELLO'
if (pass$ == 'hello')
call access_granted
else
call access_denied
endif
end
endproc