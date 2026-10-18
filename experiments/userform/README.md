# UserForm 생성·push 통합 검증

2026-10-18. 제로베이스 JSON/VBA 실험을 실제 `form build` CLI로 연결했다.
이 폴더에는 입력 fixture와 실제 Excel을 사용하는 opt-in 통합 검증만 남겼다.
이전 PowerShell 생성기/COM 직접 조작 probe는 제품 경로와 중복되므로 제거했다.

## 실행

먼저 Release 빌드를 만든다. .NET 10 + Excel + 기존 VBOM 접근 권한이 필요하다.

```powershell
dotnet build -c Release
pwsh -NoProfile -File experiments/userform/probe.ps1
# 별도 SDK를 쓸 경우:
pwsh -NoProfile -File experiments/userform/probe.ps1 -Dotnet ./_dbg/sdk10/dotnet.exe
```

probe는 정식 `vbaloop.exe`를 실행한다. `dotnet vbaloop.dll`은 자식 러너가
현재 exe를 재실행하는 기존 run/test 계약에 맞지 않으므로 사용하지 않는다.
프로젝트·워크북·결과 JSON은 git 제외 경로 `_pilot_form/<guid>/`에 보존한다.
setup은 `--no-install --skip-verify --yes`를 사용하고 보안 설정을 바꾸지 않는다.

## 검증 범위

1. CLI setup으로 빈 워크북 생성.
2. `ProbeForm.form.json` + `ProbeForm.code.vba`만으로 `form build` 실행.
3. 생성된 frm/frx와 검증용 표준 모듈을 push, `compiled` 및 저장 성공 확인.
4. CLI run으로 저장된 폼을 로드해 Frame 안 TextBox와 버튼 Click 이벤트를 확인.
5. --force 없는 재생성을 거부하고 이전 파일쌍 해시가 유지되는지 확인.
6. 잘못된 명세 + --force도 기존 쌍을 보존하는지 확인.
7. VBA 입력을 수정해 --force 재생성 → push/compile → 이벤트 결과 변경 확인.

제품 생성 경로 자체도 Export/Import 후 폼 이름·타입, 지정된 속성·코드를 검증한다.
루트 크기는 1pt, 컨트롤 위치/크기는 0.1pt 오차까지 허용한다.
지원 범위와 사용 예제는 [프로젝트 README](../../README.md) 및
[에이전트 스킬 정본](../../assets/vbaloop-skill/SKILL.md)에 있다.

## 앞선 탐색 실험에서 얻은 근거

- 연속 Export에서 frm은 같아도 frx 바이트는 달라졌다.
- 버튼 Caption만 바꿔도 frm 텍스트는 같았다. 현재 pull의 FRX 잡음 생략은
  실제 디자인 변경을 놓칠 수 있다. 이 별도 문제는 이번 생성 CLI에서 수정하지 않았다.
- 코드만 교체하면 디자인이 보존됐고, frm/frx Import는 한글·중첩 컨트롤을 보존했다.
- 원본 쌍을 Export해 둔 경우 실패한 Import 후 복원이 가능했다. 부분 Import나
  복원 자체의 실패까지 증명한 실험은 아니다. 기존 push 복구/poison 정책은 유지했다.

xlflow 검토 리비전: `011c19e9dfeba9c7925b712881b955f2b51253c7`.
[ExcelFormWriteService](https://github.com/harumiWeb/xlflow/blob/011c19e9dfeba9c7925b712881b955f2b51253c7/bridge/dotnet/src/Xlflow.ExcelBridge/Services/ExcelFormWriteService.cs)는
기존 폼 백업·부분 Import 제거·복원·복원 실패 시 poison을 구현한다.
[ADR-0072](https://github.com/harumiWeb/xlflow/blob/011c19e9dfeba9c7925b712881b955f2b51253c7/docs/adr/ADR-0072-userform-canonical-compiler-and-persisted-dimensions.md)의
현재 xlflow는 pure-Go 폼 컴파일러를 사용한다. 이번 구현은 기존 Excel COM 기반
vbaloop에 맞춰 작게 만들었으며 xlflow 코드를 복사하거나 실행하지 않았다.

## Test delta

- 고유 위험: 잘못된 명세/손실 인코딩 허용, 파일쌍 게시 실패로 기존 소스 훼손,
  생성한 폼이 실제 push/compile/event 경로에서 작동하지 않음.
- 소유자/계층: CliArgsTests 확장; 새 입력 경계 FormDefinitionTests, 파일 게시 경계
  FormSourceTests 추가. 실제 COM 변환은 기존 DESIGN.md §7의 수동 Excel lane 소관.
- 추가/대체: 단위 테스트 수집 +18개. 직접 COM을 호출하던 수동 probe와 생성기를
  CLI probe로 대체했다. CI에 Excel 실행이나 새 lane을 추가하지 않았다.
- 단위 테스트: 전체 791개 통과, 최초 실행 측정 402ms(테스트 runner 보고값).
  최종 문서·자산 포함 재실행도 791개 통과(484ms). self-contained 단일 exe 게시,
  help에 form 명령 노출, 실제 build → push(compiled) → run도 통과했다.
  CLI probe 성공 결과: `_pilot_form/1cccf40125bf4bc99eed74b5ffdef21f/result.json`.
  변경 전 동일 환경의 실행 시간을 재측정하지 않아 신뢰할 시간 증감은 미측정.
- 리소스 비용: CI의 Excel 시작 증감 0. 수동 통합 검증은 setup 1회, build 2회,
  push 2회, run 2회로 전용 Excel 총 7회. 기존 직접 COM probe 1회 대비 +6회이며,
  실제 명령별 수명·저장 경계를 검증하기 위한 비용이다. 실패 입력은 Excel 이전에 거부한다.
- coverage floor 및 테스트 의존성 변경 없음.
- 그림, MultiPage/TabStrip, 외부 ActiveX, 다른 DPI의 치수 정확성은 미검증/미지원.
  frm/frx 게시 중 프로세스 강제 종료는 두 파일 전체에 원자적이지 않다.
  중단 후 성공적인 재build를 거친 뒤 push해야 한다.
