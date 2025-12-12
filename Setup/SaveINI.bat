if exist mt2ofxini.new goto doit1
exit 1

:doit1
if exist mt2ofx.ini goto doit
ren mt2ofxini.new mt2ofx.ini
exit 0

:doit
if exist mt2ofx.prv goto doit2
goto doit3

:doit2
del mt2ofx.prv

:doit3
ren mt2ofx.ini mt2ofx.prv
ren mt2ofxini.new mt2ofx.ini
exit 0
