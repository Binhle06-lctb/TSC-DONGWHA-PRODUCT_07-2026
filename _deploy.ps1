# ====================================================
# DEPLOY SCRIPT - Dongwha Hong Sam Landing Page
# Run this in PowerShell from the project folder
# ====================================================

$ProjectPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectPath

Write-Host "`n=== STEP 1: Check tools ===" -ForegroundColor Cyan
node --version
git --version
if (-not (Get-Command vercel -ErrorAction SilentlyContinue)) {
    Write-Host "Installing Vercel CLI..." -ForegroundColor Yellow
    npm install -g vercel
}
vercel --version

Write-Host "`n=== STEP 2: Git init ===" -ForegroundColor Cyan
if (-not (Test-Path ".git")) {
    git init -b main
    git config user.email "lctbinh0006@gmail.com"
    git config user.name "Trung Son Pharma"
}

git add index.html vercel.json
git commit -m "feat: Dongwha Hong Sam landing page v1.0"

Write-Host "`n=== STEP 3: Create GitHub repo & push ===" -ForegroundColor Cyan
Write-Host "Checking GitHub CLI..." -ForegroundColor Yellow
if (Get-Command gh -ErrorAction SilentlyContinue) {
    gh repo create dongwha-hong-sam --public --source=. --remote=origin --push
    Write-Host "Pushed to GitHub!" -ForegroundColor Green
} else {
    Write-Host "" -ForegroundColor Yellow
    Write-Host "GitHub CLI (gh) not found. Do this manually:" -ForegroundColor Yellow
    Write-Host "  1. Go to https://github.com/new" -ForegroundColor White
    Write-Host "  2. Create repo named: dongwha-hong-sam" -ForegroundColor White
    Write-Host "  3. Copy the repo URL and run:" -ForegroundColor White
    Write-Host '     git remote add origin https://github.com/YOUR_USERNAME/dongwha-hong-sam.git' -ForegroundColor White
    Write-Host '     git push -u origin main' -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter after pushing to GitHub to continue with Vercel deploy"
}

Write-Host "`n=== STEP 4: Deploy to Vercel ===" -ForegroundColor Cyan
Write-Host "Logging in to Vercel (browser will open)..." -ForegroundColor Yellow
vercel login

Write-Host "Deploying to production..." -ForegroundColor Yellow
vercel --prod --yes

Write-Host "`n=== DONE! ===" -ForegroundColor Green
Write-Host "Your site is live on Vercel!" -ForegroundColor Green
