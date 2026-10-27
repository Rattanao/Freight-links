@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===== Check EDI Freight Links =====
echo วางไฟล์ MANIFEST (.xls/.xlsx) และ ENTER (.pdf) ไว้ในโฟลเดอร์ input ก่อน
echo.

rem หา Python ที่รันได้จริง (ตัว python.exe ของ Microsoft Store บางเครื่องเป็นแค่ทางลัดที่รันไม่ได้)
set PY=
py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY python -c "import sys" >nul 2>nul && set PY=python
if not defined PY (
  echo !! ไม่พบ Python - ติดตั้งจาก https://www.python.org/downloads/ แล้วติ๊ก "Add python.exe to PATH"
  pause
  exit /b 1
)

%PY% -c "import openpyxl, xlrd, pymupdf" >nul 2>nul
if errorlevel 1 (
  echo กำลังติดตั้งไลบรารีที่ต้องใช้ ครั้งแรกครั้งเดียว...
  %PY% -m pip install -r requirements.txt
)

set SHED=
set /p SHED=SHED NO. ของงานนี้ (เช่น 0141, ไม่รู้ให้กด Enter ข้าม): 
if not exist output mkdir output
if "%SHED%"=="" (
  %PY% scripts\build_edi.py --indir . --outdir output
) else (
  %PY% scripts\build_edi.py --indir . --outdir output --shed %SHED%
)
echo.
echo ผลลัพธ์อยู่ที่ output\EDI.xlsx  (ถ้าเปิด EDI.xlsx ค้างไว้ใน Excel ให้ปิดก่อนรัน)
pause
