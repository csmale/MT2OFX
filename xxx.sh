base="C:\Documents and Settings\Administrator\My Documents\My Projects\MT2OFX"
cd $base/Build
for i in *.vbs; do
    diff "$i" ..
done
