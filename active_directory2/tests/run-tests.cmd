@echo off
if "%*"=="" (
  python -munittest discover --verbose --failfast
  pause
) ELSE (
  python -munittest %*
)
