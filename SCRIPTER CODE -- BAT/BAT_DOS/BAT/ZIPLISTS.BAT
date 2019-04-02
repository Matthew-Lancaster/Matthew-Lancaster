@if %os%==DOS del c:\mylist.zip
@if %os%==DOS copy c:\mylist.txt c:\attach 
@if %os%==DOS PKZIP -R -P -EX C:\MYLIST.ZIP c:\attach\mylist.txt E:\1LOOK\1SWAPS\cd-rom\MINE\*.*
@if %os%==DOS copy c:\mylist.zip C:\MYDOCU~1
@if %os%==WIN del c:\mylist.zip
@if %os%==WIN copy c:\mylist.txt c:\attach 
@if %os%==WIN PKZIP -R -P -EX C:\MYLIST.ZIP c:\attach\mylist.txt g:\1LOOK\1SWAPS\cd-rom\MINE\*.*
@if %os%==WIN copy c:\mylist.zip e:\mydocu~1
