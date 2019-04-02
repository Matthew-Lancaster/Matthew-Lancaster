e:
cd e:\winnt\system32\config
cd f:\winnt\system32\config
attrib -h *.log
f:
attrib -h *.log
e:
xcopy . f: /m <c:\bat\win2kbak.tmp
attrib +h *.log
f:
attrib +h *.log
