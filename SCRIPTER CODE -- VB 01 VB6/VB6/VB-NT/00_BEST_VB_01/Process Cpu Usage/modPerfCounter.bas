Attribute VB_Name = "modPerfCounter"
Option Explicit

' Constants for the counters

' to see the full list, look at the registry key :
' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\
' Choose an Id langage (OO9 : US English, OOC : French) and look at
' the number in front of each counter (or object) in "Counter"

' Objects
Public Const COUNTERPERF_SYSTEM = 2
Public Const COUNTERPERF_MEMORY = 4
Public Const COUNTERPERF_PROCESS = 230
Public Const COUNTERPERF_THREAD = 232
Public Const COUNTERPERF_PROCESSOR = 238
' Counters
Public Const COUNTERPERF_PERCENTPROCESSORTIME = 6

