@echo off
REM Bam dup vao FILE NAY de chay cong cu RPA Nhap hoa don MISA.
REM File nay chi goi Chay_RPA.ps1 o cung thu muc (Windows khong cho bam dup chay
REM truc tiep file .ps1 -- day la cach lam chuan de "bam la chay").
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Chay_RPA.ps1"
