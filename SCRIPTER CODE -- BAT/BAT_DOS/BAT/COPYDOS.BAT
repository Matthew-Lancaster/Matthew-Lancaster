@echo off
verify on
same
if exist \cdrive.txt ECHO ABOUT TO BEGIN COPYING   C:.  D:
if exist \cdrive.txt XCOPY C:. D: /v /m 
if exist \edrive.txt ECHO ABOUT TO BEGIN COPYING   E:.  F:
if exist \edrive.txt XCOPY E:. F: /v /m 
if exist \gdrive.txt ECHO ABOUT TO BEGIN COPYING   G:.  J:
if exist \gdrive.txt XCOPY G:. J: /v /m 
if exist \hdrive.txt ECHO ABOUT TO BEGIN COPYING   H:.  K:
if exist \hdrive.txt XCOPY H:. K: /v /m 
if exist \idrive.txt ECHO ABOUT TO BEGIN COPYING   I:.  L:
if exist \idrive.txt XCOPY I:. L: /v /m 
