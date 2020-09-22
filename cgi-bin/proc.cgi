proc world
 print 'world'
endproc

proc hello
 print 'hello'
 call world
endproc

proc main
call hello
end
endproc