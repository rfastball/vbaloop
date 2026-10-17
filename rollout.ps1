# rollout.ps1 - vbaloop 버전업 원샷: test -> publish -> 온보딩된 전 프로젝트 init 수렴
# (재빌드는 실행 중인 exe가 스스로 못 하므로 스크립트가, 프로젝트 fan-out은 새 exe의 init --all이 담당)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

Write-Host "== [1/4] dotnet test ==" -ForegroundColor Cyan
dotnet test
if ($LASTEXITCODE -ne 0) { Write-Host "테스트 실패 - 롤아웃 중단" -ForegroundColor Red; exit 1 }

Write-Host "== [2/4] dotnet publish ==" -ForegroundColor Cyan
dotnet publish src\VbaLoop -c Release -r win-x64
if ($LASTEXITCODE -ne 0) {
    Write-Host "publish 실패 - vbaloop.exe가 실행 중이면(파일 잠금) 종료 후 재시도" -ForegroundColor Red
    exit 1
}

$exe = Join-Path $PSScriptRoot 'src\VbaLoop\bin\Release\net10.0-windows\win-x64\publish\vbaloop.exe'

# [3/4] PATH shim: 고정 폴더에 최신 exe 복사 + 사용자 PATH 1회 등록(멱등). PATH를 빌드 구성
# 경로에 결합하지 않고 롤아웃마다 최신으로 수렴시켜 에이전트 셸이 'vbaloop'를 바로 해석하게 한다
# (버그리포트 2026-07-11 #1). 복사/등록 로직은 exe 안(Core.SelfInstall)이 단일 소스다 —
# 배포본 사용자는 rollout.ps1이 없으므로 setup이 같은 일을 해야 하고, 여기서 그걸 재사용한다.
$shimDir = Join-Path $env:LOCALAPPDATA 'vbaloop\bin'
Write-Host "== [3/4] PATH shim ($shimDir) ==" -ForegroundColor Cyan
# --yes는 AccessVBOM을 승인하지 않는다(설계 §10.8) — 롤아웃 실행 자체를 개발 PC 소유자의 동의로 본다
& $exe setup --install-only --yes --enable-vbom
if ($LASTEXITCODE -ne 0) {
    Write-Host "설치 실패 - 실행 중인 vbaloop 종료 후 재시도" -ForegroundColor Red
    exit 1
}
if (($env:Path -split ';') -notcontains $shimDir) { $env:Path += ";$shimDir" }

$rev = (& $exe version | Select-Object -First 1)
Write-Host "== [4/4] init --all ($rev) ==" -ForegroundColor Cyan
& $exe init --all
$initExit = $LASTEXITCODE

# init --all은 벤더링 자산(src\LogShim.bas 등)을 재전파한다 → 각 프로젝트의 src 지문이
# 마지막 push와 어긋난다. init은 순수 파일시스템이라 push할 수 없으므로, 벤더링이 바뀐
# 경우 각 프로젝트에서 dist/pull 전에 vbaloop push가 선행돼야 한다(감사 G).
Write-Host ""
Write-Host "참고: 벤더링 자산이 바뀐 경우 각 프로젝트에서 'vbaloop push' 선행 후 dist/pull." -ForegroundColor Yellow

exit $initExit
