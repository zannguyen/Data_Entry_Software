# Chay_RPA.ps1
# ------------------------------------------------------------------
# Khoi dong cong cu RPA Nhap hoa don MISA (GUABAO ONE).
# KHONG bam dup truc tiep vao file nay -- Windows se mo bang Notepad.
# Hay bam dup vao "Chay_RPA.bat" o cung thu muc, file do se tu goi script nay.
# ------------------------------------------------------------------

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

Write-Host "=== RPA Nhap hoa don MISA - GUABAO ONE ===" -ForegroundColor Cyan
Write-Host "Thu muc lam viec: $PSScriptRoot"
Write-Host ""

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "LOI: Khong tim thay Python tren may nay." -ForegroundColor Red
    Write-Host "Cai Python tai https://www.python.org/downloads/ (nho tick 'Add python.exe to PATH')."
    Write-Host "Sau khi cai xong, dong cua so nay va bam lai Chay_RPA.bat."
    Write-Host ""
    Read-Host "Bam Enter de dong cua so nay"
    exit 1
}

$requiredFiles = @("gui_app.py", "misa_automation.py", "pdf_extractor.py", "product_matcher.py")
foreach ($f in $requiredFiles) {
    if (-not (Test-Path (Join-Path $PSScriptRoot $f))) {
        Write-Host "LOI: Thieu file '$f' trong thu muc nay." -ForegroundColor Red
        Write-Host "Kiem tra da copy DU ca 4 file chuong trinh vao cung 1 thu muc voi Chay_RPA.bat chua."
        Write-Host ""
        Read-Host "Bam Enter de dong cua so nay"
        exit 1
    }
}

Write-Host "Dang khoi dong ung dung..."
Write-Host "Windows se hien hop thoai UAC xin quyen Administrator -- BAM YES." -ForegroundColor Yellow
Write-Host "(Day la buoc bat buoc, khong the bo qua -- MISA can quyen Admin de dieu khien duoc.)"
Write-Host ""

python gui_app.py
$exitCode = $LASTEXITCODE

Start-Sleep -Seconds 2

if ($exitCode -ne 0) {
    Write-Host ""
    Write-Host "Ung dung dong voi ma loi $exitCode -- xem thong bao loi phia tren." -ForegroundColor Yellow
    Read-Host "Bam Enter de dong cua so nay"
}
