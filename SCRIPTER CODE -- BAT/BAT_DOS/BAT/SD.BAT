@echo off
big off
ndd d: e: f: g: /q
ncc/50
speedisk d: /F
speedisk e: /F
speedisk f: /F
speedisk g: /F
ncc/25
@IMAGE D:,E:,F:,G:
