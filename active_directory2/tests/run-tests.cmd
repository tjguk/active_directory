@echo off
if "%*"=="" (
  python -munittest discover
  pause
) ELSE (
  python -munittest %*
)
