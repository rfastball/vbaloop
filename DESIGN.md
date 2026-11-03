# vbaloop 설계

vbaloop의 동작 계약과 구현 근거. 사용법은 [README](README.md).

## 1. 개요

코드 에이전트가 Excel VBA 프로젝트에서 편집 → push → test → 판독 루프를 사람 개입 없이 돌게 하는
CLI. src 텍스트 파일과 워크북 VBProject를 손상 없이 왕복시키고, 실패해도 사용자 워크북과
사용자 Excel을 망가뜨리지 않는 것이 목표다.

## 2. 구조

- **C# 단일 exe가 Excel COM을 직접 소유한다.** 애드인·외부 런타임 없음.
  - 런타임 근거 1: RCW 수명을 명시적으로 관리해 좀비 EXCEL.EXE를 막는다.
  - 런타임 근거 2: STA 스레드를 직접 제어한다(메시지 필터, 자식 러너).
- **파일 포맷은 `VBComponent.Export` 호환**(VERSION/Attribute 헤더 포함). pull은 헤더를 보존·합성하고
  push는 주입 전에 헤더를 떼어낸다. Excel 수동 Import와도 호환된다.
- **src 규약**: UTF-8(BOM 없음) + CRLF + 파일 끝 개행 1개. 평면 `src/`.

## 3. 핵심 규칙

- 코드 텍스트는 `CodeModule.Lines`/`AddFromString`(BSTR, 메모리)으로만 왕복한다. 파일 기반 코드페이지
  변환이 없다.
- 예외는 폼(`.frm`/`.frx`)뿐이다. 레이아웃과 `.frx`는 파일로만 왕복된다. 이때 코드페이지는 시스템 ACP를
  동적으로 조회하고, Import 전 Remove는 제거 완료를 확인한 뒤에만 진행하며, Import 후 `<이름>1` 중복을 검사한다.
- 기존 모듈은 Remove+Import가 아니라 in-place 교체한다(중복 모듈 생성 방지).
- 시스템 ACP로 표현할 수 없는 문자는 push가 저장 전에 거부한다(VBE 내부 저장에서 `?`로 손실되므로).
- 매크로 호출은 항상 모듈 한정(`'Book'!Module.Proc`)이다. 비한정 호출은 동명 프로시저가 있으면 실패한다.
- Excel은 도구가 COM으로 직접 띄우고 PID·시작 시각·HWND로 식별한다.

## 4. 명령

### 공통 계약
- `--json` 봉투: `{command, ok, exitCode, hint, session, errors[], payload}`. `session`은 이번 명령이
  세션을 새로 띄웠을 때만 채워진다. 선택 필드 `update: {current, latest, command}`는 새 릴리스 알림이 있을 때만
  덧붙는다(§12.4). `hint`의 의미(실패 시 다음 행동)는 바뀌지 않는다.
- 종료 코드: 성공 0 / 환경 오류 3(세션 부재·stale·poison·busy·격리, Excel/COM 치명) / 그 외 2.
  실패하면 한 줄 에러와 `hint`(다음 행동)를 준다.
- `_dbg\` 파일: `session.json`(세션 메타), `recovery.json`(격리 마커), `sync.json`(지문 캐시),
  `lease.json`, `vbaloop.jsonl`(append-only 도구 로그), `backups\`.
- 세션이 필요한 명령(push/pull/test/run/diff)은 진입 시 Excel을 스스로 확보한다. 방식은 세션 모드가 정한다(아래).

### 세션 모드 (`[excel] session`)
- **ephemeral(기본)**: 명령마다 숨은 Excel을 띄워 정본을 열고, 명령이 끝나면 저장 없이 닫는다. push만
  쓰기로 열고 스스로 저장한다. pull/diff/test/run은 `ReadOnly:=True`로 연다 — 사용자가 정본을 열어 둔
  상태에서도 동작하고, 테스트가 바꾼 시트·상태는 다음 실행에 남지 않는다.
  - 근거: 에이전트 루프의 병목은 LLM 한 턴(수십 초)이고, 콜드 스타트는 명령당 약 1~1.5초다(실측). 반면 세션 Excel을
    명령 사이에 유지하면 stale·poison·busy·격리·lease 같은 상태 관리가 전부 그 결정에서 나온다. 일회용
    인스턴스는 도구 소유라 비정상 상태의 정리도 §11 규칙만 따르면 된다.
  - 살아 있는 명시 세션(`session start`·`attach`, `session.json`)이 있으면 그것을 쓴다. 정리는 `session stop`.
  - 메타는 `session.json`이 아니라 `_dbg\ephemeral\<id>.json`에 둔다. 자식 러너는 `--session-meta`로 받은 그 파일로
    재attach하고, poison 기록도 같은 파일로 간다.
  - 정상 종료: 정본을 저장 없이 닫는다 → 남은 워크북이 0개면 Quit → 300ms 안에 끝나지 않으면 신원 대조 후 kill.
    Quit한 Excel은 명령이 RCW를 쥐고 있는 동안 종료되지 않아(실측) dist 임시 Excel과 같은 처분을 쓴다. 남은
    워크북이 있으면(사용자가 숨은 인스턴스로 파일을 연 경우) kill하지 않고 창을 보이게 해 넘겨준다.
  - 비정상(poison·타임아웃·Quit 실패)은 §11.1의 1번 규칙(`RecoveryGate.Resolve`)으로 정리하거나 격리한다.
  - 하드 kill 등으로 정리되지 못한 메타는 다음 세션 명령이 mutex 안에서 sweep해 같은 규칙으로 정리한다.
    `doctor --kill-orphans`는 이 프로젝트의 살아 있는 일회용 인스턴스를 고아 후보에서 뺀다.
- **persistent**: 구 동작. auto-ensure가 세션을 채택하거나 새로 띄우고(`visible` 적용), 명령 사이에 유지한다.

### doctor / status / diff
- `doctor`: Excel·COM·AccessVBOM·oleacc·비트니스·격리·고아 Excel·워크북 잠금 점검. `--fix`는 AccessVBOM을
  켜고, `--kill-orphans`는 §11.5 기준의 고아만 정리한다. 둘 다 명시할 때만 동작한다.
- `status`: 세션 상태(none/active/poisoned/stale/busy/quarantined), src와 마지막 push의 동기화 여부.
- `diff`: 워크북 vs src 읽기 전용 비교. 모듈별 same/changed/casingOnly/onlyInWorkbook/onlyInSrc와
  첫 차이 행을 보고한다. `--module`은 워크북 쪽 본문을 덤프한다. 차이가 있어도 exit 0이다(`payload.clean`).
  ephemeral 모드는 읽기 전용 일회용 인스턴스로, persistent 모드는 살아 있는 세션에서만 동작한다(진단이 세션
  Excel을 띄우는 것 자체가 부작용이라 auto-ensure하지 않는다).
- `status`: `payload.mode`(ephemeral/persistent)를 함께 보고한다. ephemeral 모드의 세션 none은 정상이다.

### session
- `start`: 살아 있는 세션은 채택하고, 없으면 새 Excel을 띄운다. ephemeral 모드에서도 유지 세션을 만든다 —
  `session stop` 전까지 세션 명령은 이 Excel을 쓴다. 이벤트·링크 갱신·보안 대화상자를 끈 채로
  워크북을 연다. read-only로 열리면(다른 프로세스가 잠금) 거부한다.
- `attach`: 사용자가 이미 연 정본 워크북을 채택한다(§11.6).
- `stop`: 저장 후 닫는다. 도구가 열지 않은 워크북이 있으면 Excel을 종료하지 않는다. `--no-save`,
  `--discard`(§11.3).

### pull
- `CodeModule.Lines`로 본문을 읽어 src에 쓴다. 기존 파일의 헤더는 보존하고, 신규 모듈은 표준 헤더를 합성한다.
- VBE가 바꾼 식별자 대소문자는 src 철자로 되돌린다(diff 소음 방지).
- 마지막 push/pull 이후 동기화되지 않은 src 편집이 있으면 막는다. 덮어쓰려면 `--force`.
- 워크북에 없는 src 파일은 지우지 않고 경고만 한다.

### push
0. 지문 캐시: src가 마지막 push와 같으면 아무것도 하지 않는다(`--force`로 무시). Excel 확보 전에 판정한다.
1. **백업**: 첫 변경 직전에 `SaveCopyAs`로 `_dbg\backups\<UTC id>-push\`에 사본과 `metadata.json`을 남긴다.
   `[backup] keep`(기본 10, 0은 비활성)개를 보존한다.
2. 파일 검사: BOM 제거, ACP 가드, 헤더 제거.
3. 모듈별 스냅샷 후 in-place 교체. 신규 모듈은 생성, 워크북에만 있는 모듈은 `--prune`일 때만 제거한다.
   문서 모듈(Sheet/ThisWorkbook)은 교체만 가능하다.
4. 무결성 검사: 예상하지 못한 컴포넌트 이름(`<이름>1` 등)이나 빈 모듈이 생기면 실패한다.
5. **컴파일 게이트**: 저장 전에 VBE 컴파일(§9.8과 같은 `VbeCompile`)을 실행한다. 컴파일 오류·검증 불가면 저장하지
   않고 실패한다(exit 2, `payload.compile`에 모달 본문). 컴파일되지 않는 코드는 디스크 정본에 들어가지 않는다.
   - 도구가 띄운 Excel(일회용, persistent·`session start`)에서 돈다. 모달 해제는 VBE를 닫고(Excel 종료 가능),
     타임아웃은 kill로 끊기 때문에 `attach`로 채택한 사용자 Excel에서는 생략하고 사유를 `payload.compile.detail`과
     경고 줄·hint에 남긴다. 그 세션의 컴파일 오류는 첫 run/test가 모달 본문으로 보고한다.
   - 판정은 결론이 날 때까지 폴링한다(250ms 간격, 상한 60s): VBA 모달을 보면 오류, 컴파일 컨트롤 Enabled가
     false가 되면 컴파일됨. `Execute`가 즉시 반환되므로, 예전처럼 500ms 뒤에 한 번만 읽으면 컴파일 전 상태를 읽고
     오류 코드를 unverified로 통과시켜 저장했다(BizLookup 2026-10-30~31, 재현 4회 중 3회). 상한을 넘긴 경우만
     unverified로 남는다.
   - 오류 위치: 모달 본문에는 위치가 없다. VBE는 오류 지점을 선택해 두므로(구문 오류는 행 전체, 미정의 Sub 등은
     토큰) 워치독이 모달을 닫은 뒤 VBE 닫기를 미루고, 그사이 메인 스레드가 `ActiveCodePane.GetSelection`을 읽어
     `payload.compile.location`과 오류 문구에 붙인다. 컴파일 명령 `Execute`는 비동기 게시라 즉시 반환되므로 메인
     스레드가 막혀 있지 않다. VBE를 닫으면 Excel이 종료될 수 있어 닫은 뒤에는 읽을 수 없다(2026-10-31 실측).
     코드 모듈 행은 헤더를 뺀 번호라 주입 본문과 행 텍스트를 대조해 src 파일 행으로 바꾼다. dist는 사본이라 코드
     행만 보고한다. test/run은 `Application.Run`이 VBE 닫기 전에는 반환되지 않아 이 방식을 쓰지 못한다.
   - 유지 세션에서 게이트가 실패하면 원복하지 않고 poison한다. 모달 해제 뒤 Excel 상태를 믿을 수 없고, 저장 전이라
     다음 명령이 재기동하며 여는 디스크 정본은 push 이전 상태다. 생략하던 시절에는 컴파일 오류가 push를 통과해 다음
     run에서 세션을 죽였으므로, 저장 전에 멈추는 쪽이 잃는 것이 없다.
   - 탈출구: `push --skip-compile`(1회), `[push] compile = false`(프로젝트). 이번 변경과 무관한 기존 컴파일 오류가
     있는 레거시 워크북용이다.
6. 저장. 저장하지 않은 성공은 보고하지 않는다.

`--package <txt> --workbook <파일>`은 정본 대신 업무 워크북에 패키지를 적용하는 모드다(§13.3).

실패하면 이번 push가 건드린 모듈을 전부 되돌리고 검증한다. 되돌릴 수 없으면(폼 재Import, prune,
무결성 실패, 원복 실패) 세션을 poison하고 저장하지 않는다. 일회용 인스턴스는 되돌리지 않고 저장 없이 버린다
(`workbookState = discarded`). 어느 경우든 디스크의 정본은 push 이전 상태로 남는다. `payload.recovery`에 결과를 담는다.

### rollback
`--list` / `--latest` / `--backup <id>`. 파일 작업만 한다. 유지 세션 Excel이 살아 있으면 거부한다(exit 3).
현재 파일을 `pre-rollback` 백업으로 남긴 뒤 원자적으로 교체하고 해시를 대조한다. 이후 push는 전량 재주입하고,
pull은 `--force`를 요구한다.

### test / run
- `test`: 표준 모듈의 인자 없는 `Test_*` Sub를 모듈 한정으로 실행한다. ephemeral 모드는 정본을 읽기 전용
  일회용 인스턴스로 열어 실행하므로 테스트가 바꾼 시트는 저장되지 않는다(테스트 안의 `Save`는 지원하지 않는다).
- 실행은 **자식 러너 프로세스**가 맡는다. `Application.Run`은 블로킹 호출이라 같은 프로세스에서는 시간 제한을
  걸 수 없기 때문이다. 부모는 매크로가 끝날 때까지 기다리며 취소 계기(Ctrl+C, 선택 `--timeout`, 안전 상한
  `[run] maxSeconds`)와 막는 대화상자를 감시한다. 취소 모델은 §14.
- 판정: 런타임 오류 없이 반환 AND 해당 구간의 LogShim `ok:false` 레코드 0건이면 pass.
- 어서션 싱크는 프로젝트에 벤더링한 self-contained `LogShim.bas`다. `Debug.Assert`는 헤드리스에서 아무 신호도
  내지 않는다.
- VBE 오류 대화상자는 러너의 워치독이 자동으로 닫는다(§8). 세션이 죽으면 남은 테스트는 skipped로 즉시 반환한다.
- Excel 프로세스 자체의 소멸(크래시)도 세션 사망이다. 러너와 부모가 실행 전에 Excel 프로세스 핸들을 쥐고, 호출이
  크래시 의심 오류(DISP_E_EXCEPTION·RPC 단절)로 끝나면 최대 5초 소멸을 기다린다(일반 VBA 오류 0x800A9C68은 즉시
  확인만). 실측(2026-11-02): 액세스 위반은 6ms 만에 DISP_E_EXCEPTION으로 돌아오고 프로세스는 1초 안에 사라지며
  종료 코드는 0이다 — 그래서 종료 코드가 아니라 크래시 의심 오류·0 아닌 종료 코드·재기동 Excel 중 하나로 크래시를
  판정한다(징후가 없으면 "exited"). 그 테스트는 skipped가 아니라 fail(`Excel crashed while running ...`),
  `sessionDied: true`, `payload.excelCrash`, 로그 `excel_crashed`. push/pull/diff의 COM 실패도 같은 판정을 거쳐
  poison 대신 크래시로 보고한다(실측 크래시는 push의 폼 임포트 직후).
- Office 크래시 핸들러는 크래시한 Excel을 부모로 하는 새 Excel을 보이는 창으로 띄워 정본과 복구 사본을 다시 열기도
  한다(그 사이 push는 read-only로 거부). 이 Excel은 사용자에게 보이는 복구 상태라 kill하지 않고 PID만 보고한다(§11).
- `run <Module.Proc>`: 같은 러너로 단일 매크로를 실행하고 반환값을 보고한다.

### init
관리 아티팩트를 만들고 드리프트를 수렴한다. Excel이 필요 없고 멱등이다. `--check`는 검사만 한다.

| 아티팩트 | 관리 방식 |
|---|---|
| `vbaloop.toml` | 없으면 루트의 유일한 `*.xlsm`으로 스캐폴드. 사용자 파일이라 덮어쓰지 않고 `[vbaloop] managed_by` 한 줄만 관리한다(§12) |
| `src\LogShim.bas` | 정본 사본(통째 관리) |
| `CLAUDE.md`, `AGENTS.md` | `<!-- vbaloop:managed:start/end -->` 블록만 관리 |
| `.claude\skills\vbaloop\` | 스킬과 참조 문서(통째 관리) |
| `.gitignore` | `# vbaloop:managed` 블록(`_dbg/`, `dist/`) |

정본은 exe에 임베드돼 있어 설치본만으로 온보딩할 수 있다. `--all`은 프로젝트들을 일괄 수렴한다(대상은 §12.6).
exe가 프로젝트의 관리 스탬프보다 오래되면 init은 아무것도 쓰지 않고 거부한다(§12.2).

### lint / bugreport / version / setup
- `lint`: 인코딩·규약·헤더·VB_Name·frm/frx 짝 전수 검사. Excel이 필요 없다.
- `bugreport`: 결함 리포트 템플릿을 만든다. 소스 레포에서는 `docs\bugreports\`, 설치본에서는 `_dbg\bugreports\`
  + GitHub 이슈 링크.
- `version`: semver + git 리비전 + 빌드 시각(§12.1).
- `update`: 최신 릴리스로 설치본 교체 후 init --all 여부 확인(§12.5).
- `setup`: §10.

## 5. `vbaloop.toml`

```toml
[project]
name = "MyProject"            # 생략 시 워크북 이름

[excel]
workbook = "MyProject.xlsm"   # 루트 기준 상대 경로(필수)
visible = true                # persistent 세션 Excel 표시 여부(일회용 인스턴스는 항상 숨김)
session = "ephemeral"         # ephemeral(기본) | persistent

[src]
root = "src"

[push]                        # 선택
compile = true                # 저장 전 컴파일 게이트(일회용 인스턴스에서만)

[dist]                        # 선택
version = "1.2.0"
exclude = ["Test_*"]
bootstrap = "modBootstrap.Bootstrap"
stripComments = false

[backup]                      # 선택
keep = 10

[run]                         # 선택
maxSeconds = 3600             # run/test 안전 상한(§14.5). 소요시간 추정이 아니라 버려진 실행 정리용

[vbaloop]                     # init이 기록(§12.2) - 직접 수정 금지
managed_by = "0.3.0"
```

cwd에서 위로 올라가며 처음 만나는 `vbaloop.toml`의 위치가 프로젝트 루트다. 파서는 지원하는 부분집합
(문자열, bool, 정수, 문자열 배열, 주석)만 읽고, 그 밖의 문법은 조용히 무시하지 않고 오류로 거부한다.

### 폼 신규 생성 (`form build`)

`form build --spec <json> --code <vba> [--force]`는 프로젝트 mutex와 lease를 획득해
별도의 도구 소유 Excel에서 JSON 디자인을 `VBComponents.Add(3)`와 MSForms로 구성한다.
strict UTF-8 및 ACP 가드, 필드·타입·이름·치수 검증은 Excel 실행 전에 완료한다.
지원 컨트롤은 Label/TextBox/CommandButton/Frame이며 임의 ProgID/속성 호출은 허용하지 않는다.
코드는 VbaProjectOps.ReplaceCode를 재사용한다. Export/Import 후 속성·코드를 검증하고
정규 UTF-8 frm와 binary frx를 src에 게시한다. 전체 컴파일은 타 모듈 참조가 있는 폼도
지원하기 위해 기존 push 게이트에 맡긴다. build는 사용자 코드를 실행하지 않는다.

기존 쌍은 --force 없이는 덮어쓰지 않는다. 게시 중 오류는 교체한 파일을 복원하며,
복원 실패 시 `_dbg/form-build/<id>` 백업을 보존하고 push 금지 안내를 준다.
두 파일의 게시가 프로세스 강제 종료까지 원자적인 것은 아니다. 중단 뒤에는 재build 후 push한다.
JSON/VBA는 생성 입력이며 push 지문에는 생성된 frm/frx만 포함된다. pull은 생성 입력을
역갱신하지 않는다. 기존 폼 재Import 복구/poison 및 pull FRX 잡음 정책은 별도다.

## 6. 쓰지 않는 API

- `GetActiveObject`(ROT): 임의의 Excel 인스턴스를 돌려줘 대상을 특정할 수 없다. 재attach는 HWND의
  EXCEL7 창에서 `AccessibleObjectFromWindow`로 얻고 PID를 대조한다.
- `VBComponents.Remove` + `Import`로 코드 교체(폼 제외).
- 파일 기반 ANSI 변환(Export/Import, 고정 코드페이지).
- 애드인 의존(LogShim도 self-contained).

## 7. 검증

순수 로직(헤더, 인코딩 가드, toml, diff, 판정 매트릭스, 백업·격리 상태 머신 등)은 xUnit으로, Excel 경로는
실제 Excel에서 확인한다.
1. pull 결과가 기존 src와 정규화 후 일치한다
2. push → pull 왕복이 src 바이트 기준으로 무손실이다(VBE 대소문자 변경은 pull이 흡수)
3. ACP 밖 문자는 push가 저장 전에 거부한다
4. 무결성 검사가 (a) 의도치 않은 `<이름>1` 중복을 잡고 (b) 정상 신규 모듈은 통과시킨다
5. 폼 왕복과 코드비하인드의 한글 주석이 보존된다
6. 세션 수명주기(멱등 start, stale·poison 판정)
7. test의 런타임·컴파일 오류 대화상자 처리, 실행 취소(§14.7 E2E: 협조적 취소·유예 뒤 강제 경로·막는 대화상자·
   Ctrl+C 1·2회·부모 사망·실행 잠금 거부·구 LogShim 거절·test 모드 취소·중단된 매크로 뒷정리)
8. 일회용 모드: 명령 후 EXCEL 프로세스·메타 잔존 0, 타임아웃 후 정리, 하드 kill 잔재의 다음 명령 sweep,
   명시 세션 채택, persistent 회귀
9. push 컴파일 게이트: 컴파일 오류면 저장 안 함(디스크 무변경), `--skip-compile` 통과, 정상 코드 `compiled`
10. 패키지 왕복: 시트 데이터가 있는 픽스처에 dist `.txt` → `push --package` 후 셀 값·시트 수·이름 정의가 불변이고
    코드는 패키지와 일치한다. 컴파일 오류 패키지는 업무 파일을 디스크에서 바꾸지 않는다
11. 패키지 해시: 줄끝 LF 변환은 통과, 1바이트 변경(메타 포함)·꼬리 줄 잘림·CP949 재인코딩은 거부
12. 패키지 모드는 프로젝트 폴더 안에서 실행해도 `sync.json`·lease를 바꾸지 않는다

## 8. 부가 기능

- **모달 자동 회수**: 러너 안의 워치독이 VBE 대화상자를 찾아 닫는다. 버튼은 UI 언어와 무관하게 고른다.
  - 런타임 오류: 액세스 키 E(End). Debug는 절대 누르지 않는다(break 모드가 남는다).
  - 컴파일 오류: IDOK/IDCLOSE → 단일 버튼. 그 뒤 VBE를 닫아야 `Application.Run`이 반환된다.
  - 숨은 대화상자도 대상이다. 컴파일 직후에도 잠시 더 감시한다.
  - MsgBox 같은 사용자 대화상자는 닫지 않는다. 부모가 5초 넘게 남은 막는 대화상자를 감지해 실패로 끝낸다(§14.5).
- **lease**: 변경 명령이 프로젝트 lease를 자동으로 잡는다. 소유자는 세션 id 환경변수(Claude Code는
  `CLAUDE_CODE_SESSION_ID`, 생존 판정은 `CLAUDE_PID`) 또는 에이전트 호스트 프로세스(PID+시작 시각)다. 환경변수가
  우선인 이유는 부모가 먼저 끝난 백그라운드 셸에서도 계보와 달리 끊기지 않기 때문이다. 계보 탐색은 sihost 같은
  경계에서 멈춘다(Claude 데스크톱의 부모가 sihost라 모든 세션이 하나의 소유자로 뭉쳤다). 죽었거나 20분 유휴인
  소유자는 회수한다. 작업을 끝낸 세션이 열린 채 남으면 에이전트 프로세스가 살아 있어 사망 회수가 성립하지
  않으므로(Codex 스레드 id에는 PID도 없다), 관리 블록이 작업 끝의 `vbaloop lease release`를 지시하고 유휴
  회수는 그 지시를 놓친 경우의 백스톱이다. 같은 프로젝트의 동시 명령은 OS mutex로 직렬화한다.
- **lint**: push 가드의 일괄 사전 검사 버전.
- **스킬 계층**: 루프 요약은 `CLAUDE.md` 관리 블록에, 절차와 실패 플레이북은 프로젝트 로컬 스킬
  (`.claude\skills\vbaloop\SKILL.md`)에 둔다. 배포(dist)와 VBA 함정은 필요할 때만 읽는 참조 문서로 분리한다.
- **외부 애드인 참조**: Tools > References로 참조한 `.xlam`은 도구 세션에서도 자동 로드된다. pull은 대상
  워크북만 다룬다.

## 9. dist

워크북을 건드리지 않고 배포용 클린 사본을 만든다. push가 선행돼야 한다(마지막 push == src).
- **9.0 부트스트랩**: 시트 구성처럼 코드로 재생성 가능한 것은 `[dist] bootstrap` 프로시저로 둔다. base
  워크북에는 재생성할 수 없는 것(폼 그림 등)만 남긴다.
- **9.1 제외**: 모듈 헤더 `' @dist exclude` 또는 `[dist] exclude` 글롭. 모듈 단위로만 제외한다.
- **9.2 참조 검사**: 제외된 모듈의 이름·Public 프로시저를 남은 모듈이 참조하면 Excel을 띄우기 전에 거부한다.
  이름은 범위를 따져 맞춘다: 클래스 멤버는 후보가 아니고, 지역·매개변수·참조 모듈 자신의 선언이 가린 이름은 넘긴다.
- **9.3 LogShim 무해화**: 배포본에서만 `IS_RELEASE = True`로 바꾼다. 호출부는 그대로 두어 컴파일이 깨지지 않는다.
- **9.4 버전 스탬프**: 합성 모듈 `modBuildInfo`(상수)와 문서 속성(`vbaloop_BuildVersion`/`_BuildRev`/`_BuiltAt`,
  패키지 적용 판정용 `vbaloop_Target`(§13.4)). `[dist] version`은 semver여야 한다(§13.1).
- **9.5 주석 제거**: opt-in, 토크나이저 기반(문자열 안의 `'` 보존). 식별자 난독화는 하지 않는다(문자열로 결합된
  이름이 조용히 깨진다).
- **9.6 실행**: on-disk 워크북을 스테이징 파일로 복사해 별도 임시 Excel에서 변형하고, 성공하면 원자적으로
  게시한다. 실패하면 기존 배포본을 그대로 둔다.
- **9.7 형식**: `--out` 확장자로 정한다(`.xlsm`, `.xlam`, 또는 텍스트 패키지 `.txt`(§13)). 기본 위치는 `dist/`이고 init이 gitignore에 넣는다.
- **9.8 컴파일 검증**: 게시 전에 VBE 컴파일(CommandBars Id 578)을 실행한다. 컴파일 오류면 게시하지 않고
  대화상자 내용을 보고한다. `--skip-compile`로 건너뛸 수 있다.

## 10. 설치·배포

- **10.1** 대화형 진입은 사람용인 `setup`과 패키지 드래그 앤 드롭(§13.6) 둘뿐이다. 나머지는 에이전트용이라
  비대화형이고 `--json`을 지원한다.
- **10.2** `setup`: 환경 점검 → AccessVBOM(동의 후) → `%LOCALAPPDATA%\vbaloop\bin` 설치 + 사용자 PATH →
  폴더 선택 → 워크북 생성 또는 채택 → `init` → 검증(신규는 push+test 초록불, 기존은 pull).
  `--folder <경로> --yes`면 비대화형이다. exe를 더블클릭하면 setup으로 진입한다.
  - 시작점은 셋 중 하나다. 기존 워크북(`--workbook`, exe에 드래그 앤 드롭, 파일 선택), 새 프로젝트(`--new <이름>`),
    준비된 폴더(`--folder`). 인자 없는 대화형 실행은 앞의 둘 중 하나를 묻고, 비우면 폴더 선택으로 간다.
  - 셸에서 인자 없이 부르면(설치 스크립트가 이어서 부르는 경우 포함) 현재 폴더를 먼저 제안한다. 프로젝트면
    다시 세팅, `.xlsm`이 1개면 그것을, 여러 개면 번호로 고르게, `.`으로 시작하는 항목 말고 비었으면 여기에 새
    프로젝트(폴더 이름 = 워크북 이름). 거절하면 위의 일반 흐름으로 간다. 더블클릭은 cwd가 exe 위치라 건너뛴다.
  - 기존 워크북은 원본을 건드리지 않는다. 원본 옆 `<이름>\`에 사본을 만든다(읽기 전용 특성 해제). 그 이름이
    남의 폴더면 `<이름> (2)`…를 쓰고, 이전 setup의 흔적(같은 워크북을 가리키는 toml, 또는 원본과 같은 사본)이
    있으면 재사용해 재실행이 멱등이다. 원본이 이미 프로젝트 안에 있으면 그 자리에서 쓴다.
  - 새 프로젝트의 부모 폴더는 셸에서는 cwd, 더블클릭이면 `%USERPROFILE%\vbaloop-projects`다. 더블클릭의 cwd는
    exe 위치라 의미가 없고, 문서·바탕화면은 OneDrive로 리디렉션되기 쉬워 피한다.
  - 대화형이고 PATH에 claude가 있으면 끝에서 그 폴더에 `cmd /k claude` 새 창을 띄울지 묻는다. 마법사가 claude를
    자식으로 기다리지 않는 것은 Excel RCW를 쥔 채 살아 있지 않기 위해서다.
- **10.3** 릴리스 exe는 self-contained 단일 파일이다(런타임 설치 불필요). Excel COM을 늦은 바인딩으로 쓰므로
  트리밍·AOT는 하지 않는다.
- **10.4** 릴리스: `v*` 태그 → 테스트 → publish → 스모크 → zip/exe/checksums → attestation → GitHub Release.
  `-`가 붙은 태그는 prerelease다.
- **10.5** `install.ps1`: 최신 릴리스를 받아 SHA-256을 확인(실패 시 설치하지 않음)하고 설치·PATH 등록까지 한다.
  ASCII 전용이고 호출자 셸을 닫지 않는다. 다음 단계는 셸 종류로 갈린다.
  - 사람의 터미널(ConsoleHost이고 입력이 리다이렉트되지 않았으며 `-NonInteractive`·`CI`가 아님)이면 현재 위치에서
    `vbaloop setup`으로 이어진다. "워크북이 있는 폴더(또는 빈 폴더)에서 한 줄"이 곧 프로젝트 시작이 되게 하는 것이
    목적이다. 설치만 하려면 `VBALOOP_NO_SETUP=1` 또는 `-NoSetup`, 강제로 이으려면 `-Setup`.
  - 그 외(에이전트·CI)에는 프롬프트를 띄우지 않는다. 입력이 없으면 프롬프트는 멈추거나 EOF를 읽기 때문이다.
    대신 doctor를 돌리고 에이전트용 안내(exe 절대 경로, AGENT_SETUP.md, `setup --plan`, `--enable-vbom` 동의 규칙)를
    출력한다.
- **10.6** PATH를 기록한 뒤 `WM_SETTINGCHANGE`를 브로드캐스트해 새 터미널에 바로 반영한다.
- **10.7** 프로젝트에 주입하는 VBA 텍스트(LogShim, 샘플, `modBuildInfo`)는 ASCII 전용이다. 어떤 ACP에서도
  push 가드를 통과한다.
- **10.8** 에이전트 주도 설치. 사용자는 코드 에이전트에게 [AGENT_SETUP.md](AGENT_SETUP.md)를 따르라고만 하고,
  에이전트가 설치 → 계획 → 동의 → 실행 → 보고를 한다. 이를 위한 계약은 셋이다.
  - **동의 분리**: `--yes`는 프롬프트 생략일 뿐이고 AccessVBOM(보안 설정) 변경을 승인하지 않는다. 비대화형에서
    AccessVBOM이 꺼져 있으면 `--enable-vbom`이 있을 때만 켜고, 없으면 exit 3 + `payload.consentRequired`로 멈춘다.
    에이전트가 프롬프트를 넘기려고 붙인 플래그 하나로 보안 설정이 바뀌는 길을 막는 것이 목적이다.
  - **`setup --plan`**: 레지스트리·파일시스템을 읽기만 해서 프로젝트 결정, 바뀌는 것(`changes`), 동의 항목
    (`consent`: 동의/거절 시 붙일 플래그), 실행 명령(`command`)을 돌려준다. 막는 요인이 있으면 ok=false지만
    payload는 채운다. `--new`는 cwd 기준이라 `command`에는 확정된 절대 경로를 `--folder`로 넣는다.
  - **`payload.next`**: setup을 부른 에이전트 세션은 다른 폴더에서 돌고 있어 새 프로젝트의 관리 블록·스킬이
    로드돼 있지 않다. 그 세션에서 이어갈 방법(`continueHere`)과 프로젝트 폴더에서 새 세션을 여는 명령
    (`newSession`, PATH에서 찾은 claude/codex별)을 준다. 에이전트 셸의 PATH는 설치 전으로 굳어 있어 exe는 절대 경로다.

## 11. 세션 안전 모델

사용자가 세션 Excel에서 다른 작업을 할 수 있으므로, 도구는 사용자 작업을 잃을 수 있는 kill을 하지 않는다.

### 11.1 kill 규칙
Excel을 종료하는 경로는 넷뿐이다.
1. 자동 정리(타임아웃·poison·stale, 일회용 잔재 sweep): 도구가 띄운 세션 + PID·시작 시각 일치 + 세션 워크북 외에
   열린 워크북이 없음을 확인한 경우 일회용 인스턴스는 바빠도 창 제목으로 확인할 수 있다(§14.3).
2. `session stop --discard`(명시): 도구가 띄운 세션 + 신원 일치 + 다른 워크북이 보이지 않는 경우.
3. `doctor --kill-orphans`(명시): §11.5의 고아.
4. 일회용 인스턴스의 정상 종료(§4 세션 모드): 이 명령이 띄운 인스턴스 + 정본을 닫은 뒤 열린 워크북 0개 확인 +
   Quit 송신 + 신원 일치.

그 밖의 경우는 모두 격리다.

### 11.2 프로브
응답하지 않는 Excel에 대한 COM 호출은 무기한 블록될 수 있다. 그래서 워크북 목록 확인은 자식 프로세스
(`__probe`)가 8초 제한으로 수행한다. 결과는 Responsive / Busy / Unresponsive / Windowless 중 하나다.
바쁜 일회용 Excel은 COM 대신 창 제목으로 판독하는 폴백을 둔다(§14.3).

### 11.3 격리
kill할 수 없으면 `_dbg\recovery.json`을 남긴다. 이후 세션 명령은 exit 3과 해제 안내를 낸다.
- `recovery clear`: 기록된 프로세스가 사라졌을 때 해제한다(`--force`는 강제).
- `session stop --discard`: §11.1의 2번 규칙으로 정리하고 해제한다.
- 에이전트가 현재 작업의 격리된 테스트·측정용으로 직접 시작했고, PID·시작 시각이 기동 기록과 일치하며,
  미저장 변경이 폐기 가능한 작업 산출물뿐이고 사용자에게 넘기거나 다른 워크북 작업에 쓰지 않은 세션은
  추가 확인 없이 `session stop --discard`로 정리하고 계속한다. Busy/Unresponsive여도 이 근거는 유효하지만,
  `owner = tool`이나 숨은 창만으로는 충분하지 않다. 근거가 불명확하면 사용자에게 묻되 기존 폐기 승인은 재사용한다.
  사용자 attach·Office 복구 Excel은 예외 대상이 아니며, discard 거부를 직접 kill·마커 삭제·고아 정리로 우회하지 않는다.
- 기록된 프로세스가 사라지면 다음 명령이 마커를 지우고 진행한다.
- 이때 이 프로세스는 응답 없는 Excel의 RCW 정리를 건너뛰고 결과 출력 직후 종료한다(정리 호출이 무기한 블록되므로).

### 11.4 busy
셀 편집이나 대화상자로 Excel이 호출을 거절하면 약 3초 재시도한 뒤 Busy(exit 3, 재시도 가능)로 보고한다.
stale로 취급하지 않는다.

### 11.5 고아 정리
창이 없고, 현재 Windows 세션이며, 명령줄에 `/automation`이 있고, 격리 중이 아니며, 현재 세션도 아닌 Excel만
대상이다. Word·PowerPoint에 임베드된 Excel(`-Embedding`만 있음)과 다른 사용자 세션은 보고만 한다.
도구가 띄운 PID는 `%LOCALAPPDATA%\vbaloop\excel-pids.json`에 기록해 표시한다.

### 11.6 attach
사용자 Excel에서 정본 워크북을 찾아(EXCEL 창별 NATIVEOM, 경로 비교) `owner = external`로 채택한다.
- push/pull/test/run은 그 Excel에서 동작한다. 사용자 Excel의 설정은 바꾸지 않는다.
- `stop`은 분리만 한다.
- poison·타임아웃은 항상 격리한다.
- read-only로 열려 있거나 여러 인스턴스에 열려 있으면 거부한다.

## 12. 버전 드리프트 방지

exe는 PC마다 한 번 설치되고(`%LOCALAPPDATA%\vbaloop\bin`), 관리 파일은 프로젝트 레포마다 벤더링된다.
그래서 "이 PC의 exe"와 "이 레포의 관리 파일"의 버전이 어긋날 수 있다. 정확성은 로컬·오프라인 판정(스탬프)이
맡고, 네트워크 업데이트 알림은 편의로 그 뒤에 둔다. 스탬프 거부가 이미 갱신을 지시하므로 필수 갱신에는
네트워크가 필요 없다.

### 12.1 exe 버전
- 릴리스는 태그에서 `-p:Version`을 주입한다. 로컬·CI 빌드는 `StampGitRevision`이 마지막 `v*` 태그에서
  유도한다(`v0.3.0` + 5커밋 = 버전 `0.3.0`). 커밋 수와 `-dirty`는 리비전(`v0.3.0-5-gabc1234-dirty`)에만
  남는다. git이나 semver 태그가 없으면 `0.1.0` / `nogit`.
- repo 빌드 exe가 `0.1.0`으로 비교되면 모든 프로젝트에서 거부되거나 스탬프를 끌어내리므로, 비교 단위는 반드시
  태그 기반 semver다. 비교는 semver 2.0 우선순위(prerelease < 같은 번호의 정식, 빌드 메타데이터 무시).

### 12.2 관리 스탬프 `[vbaloop] managed_by`
관리 파일(CLAUDE.md/AGENTS.md 관리 블록, `.claude\skills\vbaloop\`, `.gitignore` 블록, `src\LogShim.bas`)을
마지막으로 쓴 exe 버전. init(과 init을 부르는 setup)만 기록한다. toml은 사용자 파일이라 이 한 줄만 라인 단위로
upsert하고 나머지는 보존한다.

| exe vs 스탬프 | init | init --check |
|---|---|---|
| 스탬프 없음 | 수렴 + 스탬프 기록 | toml을 드리프트로 보고(exit 2) |
| exe = 스탬프 | 수렴 | 일반 검사 |
| exe > 스탬프 | 수렴. 관리 파일을 실제로 쓸 때만 스탬프를 올린다 | 일반 검사(관리 파일이 같으면 ok) |
| exe < 스탬프 | **아무것도 쓰지 않음**, exit 3, hint `vbaloop update` | `state: "exe-older"`, exit 3 |
| 스탬프가 semver 아님 | 거부, exit 2 | `state: "invalid"`, exit 2 |

- 스탬프는 내려가지 않는다. 더 오래된 exe는 쓰기 자체를 거부하고, 쓰는 경로도 낮은 값으로는 기록하지 않는다.
- exe가 더 새로워도 관리 파일이 같으면 toml을 건드리지 않는다. exe 갱신만으로 모든 레포에 diff가 생기는
  churn을 피하기 위해서다. 같은 정본을 가진 사이 버전 exe는 그대로 수렴할 수 있고, 정본이 바뀌는 순간
  스탬프가 오른다.
- 스탬프는 목록의 첫 항목으로 먼저 기록된다. 중간에 쓰기가 실패해도 스탬프가 파일보다 낮게 남지 않는다.
- exe-older에서는 아티팩트를 비교하지 않는다(`artifacts: []`). 더 새 정본과의 바이트 차이는 드리프트가 아니고,
  그것을 고치라는 신호가 곧 다운그레이드이기 때문이다. 관리 스킬과 CLAUDE.md 블록은 에이전트에게 이 상태를
  init으로 고치지 말고(스탬프 수정·삭제 포함) 사용자에게 `vbaloop update`를 요청하라고 지시한다.
- `init --all`도 프로젝트마다 같은 규칙이다. 거부된 프로젝트는 `[REFUSED]`로 보고하고 나머지는 계속 수렴한다.
  하나라도 거부되면 exit 3.
- 스탬프 판독은 TomlMini가 아니라 라인 스캔이다. 미래 exe가 이 exe가 모르는 toml 문법·섹션을 써도 스탬프는
  읽혀야 "exe가 오래됨"을 정확히 말할 수 있다. 같은 이유로 `[vbaloop]` 섹션의 미지 키는 거부하지 않고, 다른
  섹션의 미지 키·섹션이라도 스탬프가 exe보다 새로우면 오류를 exe-older(exit 3, `vbaloop update`)로 바꿔 보고한다.
- 스탬프 도입 이전 exe(`[vbaloop]` 섹션을 모름)는 스탬프된 프로젝트에서 모든 명령이 "미지 섹션"으로 실패한다.
  이 exe들은 init도 실패하므로 스탬프된 프로젝트를 다운그레이드할 수 없다. 해결은 exe 갱신이다.

### 12.3 다른 명령의 경고
push/test/pull/run 등 프로젝트를 읽는 명령은 exe-older여도 막지 않고 stderr에 경고 한 줄만 낸다
(`--json`에서도 stdout 봉투는 그대로). 근거: 이 명령들은 관리 파일을 쓰지 않아 다운그레이드를 일으키지 않고,
새 exe는 옛 LogShim 레코드를 계속 읽는다(LogShim 형식은 이번 변경에서 바뀌지 않았다 - 형식을 바꿀 때는
LogShim에 버전 상수를 두고 test가 옛 워크북 LogShim을 감지해 "push 필요"로 멈춰야 한다). 루프를 세울 이유는
없지만 사람이 알아야 하므로 경고한다. `status`는 `managed`(by/exe/state) 필드와 한 줄을, `doctor`는
`project.managedBy` 점검(exe-older면 FAIL, hint `vbaloop update`)을 보여 준다.

### 12.4 새 릴리스 알림
- 명령은 캐시 `%LOCALAPPDATA%\vbaloop\update.json`(`checkedAt`, `latest`, 실패 시 `error`)만 읽는다. 캐시가
  24시간보다 오래됐으면 분리된 자식 `vbaloop __update-check`를 띄우고 바로 진행한다. 자식은 GitHub
  `releases/latest`(rfastball/vbaloop)를 3초 제한으로 조회해 캐시를 쓴다. 명령 자체에는 지연이 없고, 결과는 다음
  명령부터 보인다.
- 자식은 핸들을 상속하지 않고 창 없이 뜬다(CreateProcessW, bInheritHandles=false). `Process.Start`는 항상 핸들을
  상속시켜, 호출자가 읽는 stdout 파이프가 자식이 끝날 때까지 열린 채로 남기 때문이다.
- 조회가 실패해도(오프라인, 정식 릴리스 없음 404) `checkedAt`을 기록한다. 명령마다 자식이 뜨지 않게 하기 위해서다.
  동시에 뜬 자식은 뮤텍스로 하나만 조회한다.
- 건너뛰는 경우: repo 빌드 exe(RepoLocator가 VbaLoop.sln을 찾음, 갱신 경로는 rollout.ps1), CI(`CI`,
  `GITHUB_ACTIONS`, `TF_BUILD`), `VBALOOP_NO_UPDATE_CHECK=1`.
- 알림은 `status`, `doctor`, `version`, `init --check`에서만 낸다. 캐시 확인·자식 기동도 이 명령들에서만 한다.
  편집 루프의 hot path(push/test/run)는 그대로 둔다. 에이전트 루프는 시작할 때 `init --check`를 부르므로
  알림이 전달된다.
- 판정: `latest > current`. 단 정식 버전 사용자에게 prerelease는 알리지 않는다(releases/latest가 원래 주지 않지만
  캐시를 손댄 경우까지 막는다). prerelease 사용자에게 같은 번호의 정식 릴리스는 알린다.
- 출력: `--json`이면 봉투의 `update` 필드, 사람용이면 stderr 한 줄. stdout을 파싱하는 호출자를 깨지 않는다.
- 필수 갱신은 이 알림에 의존하지 않는다. exe가 오래되면 §12.2 스탬프가 오프라인으로 init을 막고 `vbaloop update`를
  지시한다.

### 12.5 `vbaloop update`
- install.ps1과 같은 계약이다. `releases/latest`의 `vbaloop.exe`를 받아 같은 릴리스의 `checksums.txt`의 SHA-256과
  대조하고, 자산이나 체크섬 항목이 없거나 값이 다르면 설치하지 않는다(fail closed).
- 대상은 설치본 `%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe`다. 실행 중일 수 있으므로 지우지 않고 `vbaloop.exe.old`로
  이름을 바꾼 뒤 새 exe를 복사한다(이전 `.old`가 아직 실행 중이면 `vbaloop.exe.<8hex>.old`). 복사에 실패하면 원래
  이름으로 되돌린다. `.old`는 설치 폴더에서 기동하는 다음 명령이 best-effort로 지운다. PATH 등록도 확인한다.
- 비교 기준은 설치본 버전이다. 실행 exe가 설치본이면 자기 버전을, 아니면 설치본에 `version --json`을 물어 쓴다.
  최신과 같거나 높으면 아무것도 하지 않는다(`--force`는 그래도 설치). `--check`는 확인만 한다. 정식 릴리스가 없으면
  (404) "갱신할 것이 없음"으로 성공한다.
- 교체 뒤 관리 파일 수렴은 **새 exe**가 한다. 이 프로세스의 임베디드 정본은 옛 버전이기 때문이다. 대화형이면
  `init --all`을 실행할지 묻고, `--yes`면 묻지 않고 실행한다. 비대화형이거나 `--json`인데 `--yes`가 없으면 실행하지
  않고 `payload.next`로 안내한다. `--json`이면 자식의 봉투를 `payload.initAll`에 담는다.
- 성공한 조회는 알림 캐시도 갱신한다.

### 12.6 프로젝트 레지스트리와 `init --all`
- 설치본 exe는 소스 레포 옆에 있지 않아 형제 탐색으로 프로젝트를 찾을 수 없다. 그래서 사용자 단위 목록
  `%LOCALAPPDATA%\vbaloop\projects.json`(`projects[]: {root, addedAt}`)을 둔다.
- 프로젝트 설정을 읽는 모든 명령(ProjectConfig.Load 성공, init의 스캐폴드 포함)이 그 루트가 없을 때만 추가한다.
  이미 있으면 파일을 쓰지 않는다. 쓰기는 네임드 뮤텍스(`Local\vbaloop-project-registry`)로 직렬화하고 임시 파일 +
  원자적 이동으로 기록한다(excel-pids.json과 같은 방식). 실패해도 명령은 계속된다.
- `init --all` 대상: exe가 소스 레포 아래에 있으면 형제 toml 프로젝트 + `_pilot_form`(기존 동작) ∪ 레지스트리,
  설치본이면 레지스트리만. 같은 루트는 대소문자·끝 구분자를 무시하고 하나로 합친다. 대상이 없으면 exit 2.
- 폴더나 `vbaloop.toml`이 사라진 레지스트리 항목은 `[GONE]`으로 보고하고 지운다(`--check`는 보고만).
  `payload.gone`.
- LogShim.bas가 바뀐 프로젝트는 워크북에 반영하려면 각자 `vbaloop push`가 필요하다. 이름을 모아 보고한다
  (`payload.needsPush`). init은 Excel 없이 파일만 다루므로 push를 대신하지 않는다.
- rollout.ps1(test → publish → `setup --install-only` → `init --all`)은 repo 빌드 exe로 돌기 때문에 형제 탐색이
  그대로 동작하고, 레지스트리의 다른 위치 프로젝트까지 함께 수렴한다.

## 13. 패키지 배포

dist가 만든 텍스트 패키지 하나로, 개발 환경이 없는 PC(폐쇄망 포함)의 업무 워크북에서 VBA 코드만 교체한다.
시트·데이터·이름 정의·서식은 건드리지 않는다. 반입물이 텍스트 파일 하나라 망연계를 통과하기 쉽고, 적용은
push 엔진을 그대로 써서 백업·in-place 교체·무결성 검사·컴파일 게이트를 공유한다. 새 명령은 없다. dist와 push에
형식 하나와 옵션 하나를 더할 뿐이라, 드물게 쓰는 경로가 따로 낡지 않는다.

### 13.1 만들기: `dist --out <파일>.txt`
- §9.7의 확장자 규칙에 `.txt`(패키지)를 더한다. 새 플래그는 없다.
- 파이프라인은 워크북 dist와 같다: push 선행 계약, 제외(§9.1), 참조 검사(§9.2), LogShim 무해화(§9.3),
  스탬프(§9.4), 주석 제거(§9.5), 컴파일 검증(§9.8). 게시 단계만 워크북 대신 패키지 텍스트를 쓴다. 컴파일
  검증을 통과하지 못하면 게시하지 않는다.
- 담는 내용은 스테이징 사본에 주입한 것과 같은 텍스트(src 헤더 + 변형된 본문, 합성 `modBuildInfo` 포함)다.
  폼은 src의 `.frm`과 `.frx`(base64)를 담는다.
- `[dist] version`이 없으면 거부한다. 적용 쪽 버전 판정(§13.4)이 이 값에 기대기 때문이다.
- `[dist] version`이 semver(`Core/SemVer`, 앞의 `v` 허용)가 아니면 출력 형식과 관계없이 거부한다(exit 2).
  워크북 dist가 남긴 스탬프도 §13.4의 비교 대상이라, 순서를 매길 수 없는 값이 업무 파일에 들어가면 안 된다.
- `--bootstrap`을 `.txt`와 함께 주면 거부하고, `[dist] bootstrap`은 경고와 함께 무시한다. 사본의 시트를 만드는
  용도라 업무 워크북에서 돌면 데이터를 덮을 수 있다. 업무 워크북에서 도는 훅은 §13.8이다.

### 13.2 포맷
```
@@VBALOOP_PACKAGE
FormatVersion=1
Target=지출관리
Version=1.3.0
Rev=8acb552
BuiltAt=2026-11-02T03:00:00Z
Modules=modMain.bas,Sheet1.cls,ThisWorkbook.cls,frmInput.frm
@@END_META

@@FILE name=modMain.bas kind=std encoding=utf-8
(src 규약 텍스트)
@@END_FILE

@@FILE name=ThisWorkbook.cls kind=workbook encoding=utf-8
(src 규약 텍스트)
@@END_FILE

@@FILE name=frmInput.frx encoding=base64
(76자 줄바꿈 base64)
@@END_FILE

@@END_PACKAGE sha256=<이 줄 앞까지 전체의 해시>
```
- 파일은 UTF-8(BOM 없음) + CRLF. 모듈 텍스트가 src 규약을 따르므로 ACP 밖 문자가 없다.
- `Target` 기본값은 `[project] name`이다.
- 해시는 하나다. 첫 바이트부터 꼬리 줄 `@@END_PACKAGE` 직전까지를 정규화(BOM 제거, 줄끝 CRLF 통일)해 계산한다.
  - 메타까지 덮는다. `Version`·`Target`을 손으로 고치면 버전 판정(§13.4)을 우회할 수 있기 때문이다.
  - 꼬리에 두면 해시 줄 자신을 빼고 계산하는 예외가 없고, 운반 중 잘린 파일은 꼬리 줄이 없어 그대로 잡힌다.
  - 파일별 해시는 두지 않는다. 전체 해시로 무결성은 충분하고, 깨진 위치는 블록 구조 검사가 알려 준다.
  - 망연계가 줄끝만 바꾼 경우는 통과시키고 내용 변경은 잡는다. 엄격 UTF-8로 읽히지 않으면 인코딩 변환(CP949
    재인코딩 등)으로 보고 거부한다.
  - 운반 중 손상과 실수로 인한 편집을 잡는 무결성 검사이고, 위변조 방지(서명)는 아니다. 누구나 다시 계산할 수 있다.
- 모듈 블록의 `kind`는 `std`/`class`/`form`/`document`/`workbook` 중 하나이고, 확장자와 맞아야 한다(`workbook`은
  하나만). `.cls` 헤더만으로는 시트·통합 문서 모듈과 클래스를 가를 수 없어, dist가 스테이징 사본의 컴포넌트
  타입과 `Workbook.CodeName`으로 정해 싣는다. `.frx` 블록에는 `kind`가 없다.
- `Modules`는 mirror 판정의 기준이다. `@@FILE` 블록과 목록이 어긋나면 거부한다.
- 모듈 본문의 행이 예약 마커와 같으면 dist가 거부한다.
- 읽는 쪽이 모르는 `FormatVersion`이면 추측하지 않고 거부한다. hint는 exe 갱신이다(반입 환경이면 새 exe 반입).

### 13.3 적용: `push --package <txt> --workbook <파일>`
- 둘은 짝이다. 하나만 있으면 exit 2. `--workbook`의 기본값을 정본으로 두지 않는 것은 패키지를 개발 정본에,
  src를 업무 파일에 잘못 넣는 사고를 막기 위해서다.
- 대상은 매크로 사용 형식(`.xlsm`/`.xlsb`/`.xlam`)만 받는다. `.xlsx`에 넣고 저장하면 코드가 사라진다.
- 프로젝트 없이 돈다. `vbaloop.toml`을 찾지 않고(`--config`는 거부), lease·`_dbg\`·지문 캐시·레지스트리를
  쓰지 않는다. 프로젝트 폴더 안에서 실행해도 그 프로젝트의 동기화 상태는 바뀌지 않는다(바뀌면 dist가
  "마지막 push == src"로 오판한다). 같은 대상 경로의 동시 실행은 경로 기반 네임드 뮤텍스로 직렬화한다.
- Excel은 항상 일회용 인스턴스다(§4 ephemeral 처분). 명시 세션과 attach는 쓰지 않는다. attach는 컴파일
  게이트를 생략하기 때문이다. 기존 열기 경로대로 이벤트를 끄고 연다(`Workbook_Open` 미실행).
- 쓰기 가능 여부는 Excel을 띄우기 전에 판정한다. `FileShare.None` 쓰기 열기를 시도해 보고 바로 닫는다. 띄운 뒤
  read-only로 열린 것을 알아채면 일회용 인스턴스를 헛되이 띄운 셈이다.
  - 다른 프로세스가 열고 있으면 exit 3이다. 닫으면 풀리는 환경 상태라 §4의 busy와 같은 분류다. hint는 "파일을
    닫고 다시 실행"이고, 같은 폴더의 `~$<파일 이름>` 잠금 파일에서 연 사람을 읽을 수 있으면 함께 싣는다.
  - 파일에 읽기 전용 속성이 있으면 exit 2다. 닫아도 풀리지 않으므로 속성 해제를 hint로 준다.
- 순서:
  1. 패키지 파싱·해시 검증, ACP 가드(Excel 전). 대상 PC의 ACP가 개발 PC와 다를 수 있어, 모듈 텍스트가 대상
     ACP로 표현되지 않으면 주입 전에 거부한다.
  2. 워크북을 열어 문서 속성을 읽고 대상·버전을 판정한다(§13.4).
  3. `--plan`이면 판정과 모듈 계획(추가·교체·제거·미변경 문서 모듈)을 보고하고 저장 없이 닫는다. 적용 기록
     (`apply.jsonl`, §13.5)도 남기지 않는다.
  4. 백업(§13.5).
  5. 주입: push 3~4단계와 같다. 패키지에 없는 표준·클래스·폼 모듈은 제거한다(`--prune` 상시). 폼은 항상
     재Import한다(비교할 근거가 없다). 문서 모듈 규칙:
     - 통합 문서 모듈은 이름이 아니라 종류로 맞춘다. 유일하고, CodeName이 Excel UI 언어에 따라
       `ThisWorkbook`/`현재_통합_문서`로 다르기 때문이다.
     - 시트 모듈은 CodeName으로 맞춘다. 패키지에 있는데 워크북에 없으면 주입 전에 거부한다(만들 수 없다).
     - 워크북에만 있는 시트 모듈은 코드가 있어도 건드리지 않고 경고한다. 현장에서 추가한 시트는 데이터
       영역이다. 그 코드가 새 코드와 맞지 않으면 컴파일 게이트가 잡는다.
  6. 컴파일 게이트(항상).
  7. 문서 속성 갱신: `vbaloop_Target`, `vbaloop_BuildVersion`/`_BuildRev`/`_BuiltAt`, `vbaloop_PackageSha256`,
     `vbaloop_AppliedAt`.
  8. 저장(원래 형식 유지). 실패하면 저장 없이 버린다. 디스크의 업무 파일은 적용 전 그대로 남는다.
- 플래그: `--plan`, `--adopt`, `--allow-downgrade`, `--json`만 받는다. `--force`(지문 캐시가 없다)와
  `--skip-compile`(코드 전체를 교체하므로 컴파일 실패는 패키지 결함이다, 개발 PC에서 고친다)은 거부한다.
  `--prune`·`--force-form-import`는 상시 켜져 있어 지정해도 효과가 없다.
- 종료 코드는 §4 공통 계약을 따른다. 판정 거부·해시 불일치·컴파일 오류는 2, 파일 잠금·Excel/COM 문제는 3이다.

### 13.4 대상·버전 판정
| 워크북 상태 | 결과 |
|---|---|
| `vbaloop_Target` 없음 | 거부. `--adopt`면 진행하고 Target을 기록한다(dist로 만든 적 없는 기존 업무 파일의 첫 적용) |
| Target 불일치 | 항상 거부. 덮어쓰는 플래그는 두지 않는다. 필요하면 패키지를 다시 만든다 |
| 패키지 버전 < 워크북 버전 | 거부. `--allow-downgrade`면 진행 |
| 워크북 버전이 semver 아님 | 거부. `--allow-downgrade`면 진행. "없음"으로 보면 다운그레이드가 조용히 통과한다 |
| 같은 버전 | 아래 규칙 |
| 패키지 버전 > 워크북 버전, 또는 워크북 버전 없음 | 진행 |

- 같은 버전이면 워크북 코드를 패키지와 비교한다. "코드가 같다"는 다음 둘을 모두 만족하는 것이다.
  - 컴포넌트 집합(이름과 종류)이 같다. 워크북에만 있는 표준·클래스·폼 모듈은 prune 대상이라 "다름"으로 센다.
    워크북에만 있는 시트 모듈은 적용이 건드리지 않으므로(§13.3) 비교에서 뺀다.
  - 모든 모듈의 코드가 같다(push의 변경 판정과 같은 정규화). 폼은 코드비하인드만 `CodeModule.Lines`로 비교한다.
  - 합성 `modBuildInfo`는 비교에서 뺀다. 빌드마다 `BuiltAt`이 달라, 같은 src를 다시 빌드한 패키지(예: 워크북 dist와
    같은 버전의 패키지)를 "다른 코드"로 거부하게 된다. 버전 자체는 앞 단계가 문서 속성으로 판정한다.
  - 폼 레이아웃은 비교하지 않는다. 연속 Export에서도 frx 바이트가 달라지고, Caption을 바꿔도 frm 텍스트가
    같아서(experiments/userform) 디자이너 비교를 믿을 수 없다. 현장에서 레이아웃만 바뀐 파일은 "이미 적용됨"으로
    지나가므로, 바로잡으려면 버전을 올려 다시 배포한다. 실제로 적용하는 경로는 폼을 항상 재Import하므로 영향이 없다.
  - 코드가 같으면 "이미 적용됨"으로 성공(exit 0)하고 백업도 저장도 하지 않는다.
  - 코드가 다르고 워크북의 `vbaloop_PackageSha256`이 이 패키지와 같으면 현장에서 코드가 바뀐 것이다. 복구로
    재적용하고, plan과 결과에 바뀐 모듈을 표시한다.
  - 코드가 다르고 해시가 다르거나 없으면 거부한다. 같은 버전 번호에 다른 코드가 실린 것이라(버전 올리기 누락)
    어느 쪽이 맞는지 판단할 근거가 없다. 덮어쓰는 플래그는 두지 않는다. hint는 `[dist] version`을 올려 다시
    만들라는 것이다.
- 버전은 `Core/SemVer`로 비교한다(prerelease 순서 반영). `+build`만 다른 두 버전은 같은 버전이라 위 규칙으로 간다.
  판정은 순수 로직이라 xUnit으로 매트릭스 전체를 검증한다.
- 워크북 dist(§9.4)도 `vbaloop_Target`을 기록해, dist 배포본에서 시작한 업무 파일은 `--adopt`가 필요 없다.

### 13.5 백업
- 4단계에서 `%LOCALAPPDATA%\vbaloop\backups\<파일 이름>-<경로 해시 8자>\<UTC id>-package\`에 사본과
  `metadata.json`(원래 전체 경로, 패키지 해시, 이전·새 버전, 대상)을 남긴다. 슬롯 형식은 `BackupStore`와 같다.
  - 사본은 `SaveCopyAs`가 아니라 디스크 파일의 바이트 복사다. 이 시점까지 아무것도 바꾸지 않았으므로 디스크
    파일이 곧 적용 전 상태이고, 바이트가 같은 사본이 복원에 더 정확하다. Excel의 쓰기 경로에도 기대지 않는다
    (E2E에서 `SaveCopyAs`가 `%LOCALAPPDATA%` 아래 새로 만든 폴더에 1004로 실패한 환경이 있었다).
  - 업무 파일 옆에 두지 않는 것은 공유 드라이브에 백업이 쌓이거나 쓰기 권한이 없어 적용 자체가 막히는 것을
    피하기 위해서다. 경로 해시는 이름이 같은 다른 폴더의 파일을 구분한다.
  - 찾기 어려워지는 대가는 출력으로 메운다. plan·결과·드래그 창과 `payload.backup`에 백업 전체 경로를 항상 싣는다.
- 자동으로 지우지 않는다. 업무 데이터가 든 파일이라 정리는 사람이 한다.
- 복원은 사본을 원래 위치로 복사해 쓰면 된다. `rollback`은 프로젝트 명령이라 이 백업을 다루지 않는다.
- 같은 백업 폴더(`<파일 이름>-<경로 해시 8자>\`)의 `apply.jsonl`에 적용 기록(시각, 원래 경로, 버전, 결과,
  백업 id)을 덧붙인다. 업무 파일 옆에는 아무것도 쓰지 않는다는 위 원칙을 따르고, 같은 업무 파일의 이력과 사본이
  한 폴더에 모인다. 결과와 `payload.backup`에 이 경로도 싣는다.
- 비영구 VDI처럼 로그오프 때 프로필이 초기화되는 환경에서는 백업도 사라진다. 그런 환경이면 결과에 나온
  경로에서 사본을 따로 옮겨 둬야 한다(도구는 환경을 판별하지 않는다).

### 13.6 드래그 앤 드롭
- exe 위에 `.txt`와 워크북을 함께 놓으면 패키지 모드로 들어간다. `.txt`만 놓으면 워크북 선택 창을 띄운다.
  인자에 `.txt`가 없으면 기존 setup 진입(§10.2)이다.
- plan 요약을 보여 주고 Y/N을 물은 뒤 적용한다. 결과를 보여 준 창은 키 입력까지 유지한다.
- CLI `push --package`는 다른 명령처럼 비대화형이다. 확인을 묻는 것은 드래그 진입뿐이다(§10.1).

### 13.7 오프라인 환경
- 릴리스 exe는 self-contained라(§10.3) 런타임 반입이 필요 없다. 설치는 exe 복사로 충분하다.
- 새 릴리스 알림(§12.4)은 조회 실패를 기록하고 넘어가며, push 경로에서는 아예 조회하지 않는다.
- 대상 PC에도 AccessVBOM이 필요하다. 꺼져 있으면 기존 진단 hint(`setup --enable-vbom` 동의 흐름)를 준다.

### 13.8 적용 후 훅 (예약, 1차 범위 밖)
- `[dist] postUpdate = "Module.Proc"` → 패키지 `PostUpdate=` 헤더. 적용 시 컴파일 게이트 뒤·저장 전에 모듈
  한정으로 실행한다. 모달·타임아웃은 dist bootstrap 규칙을 따르고, 실패하면 저장하지 않는다.
- 용도는 시트 구조 마이그레이션(열·시트 추가)이다. 1차 구현은 포맷 키만 예약하고, 읽는 쪽은 이 헤더가 있으면
  "지원하지 않는 훅"으로 거부한다(조용히 건너뛰면 마이그레이션 없이 새 코드만 들어간다).

### 13.9 GUI 업데이터
GUI 업데이터의 화면 규약은 `src/VbaLoop.Updater/DESIGN.md`에 둔다. 배포는 `build-updater.ps1`이
같은 소스의 CLI를 먼저 빌드해 해시 manifest와 함께 GUI에 내장한다. `__package-ui`는 GUI 전용 내부
프로토콜(version 1)이며 stdout 최종 JSON, stderr 단계 JSON을 사용한다. 일반 CLI 봉투는 바뀌지 않는다.
GUI 적용은 검사 때의 패키지/워크북 SHA-256을 요구하고 열기 전·후 워크북 변경을 검사한다.
`Unverified` 컴파일은 GUI에서 저장을 막으며 일반 CLI 정책은 유지한다. 저장 시작 뒤 결과를 잃으면
`unknown`, 저장 확인 뒤 정리 실패면 `saved` 상태로 구분한다.

GUI 전용 워치독은 저장 전 단계에서 2분간 진행이 없으면 신원이 확인된 도구 Excel만 정리하고 자식을 종료한다.
저장 진입과 타임아웃은 같은 잠금으로 결정한다. 저장 중에는 강제 종료하지 않고, 2분 지연 뒤 사용자가
창을 닫고 백그라운드 저장을 계속할 수 있게 한다. GUI 경로는 결과 출력 뒤 RCW 종료 대기 없이 자식을 종료한다.

검증(2026-11-04): 내장 CLI 단일 exe 스모크 통과. 실제 GUI에서 샘플 1.2.0 → 1.3.0 적용 후
셀 데이터·서식·시트 구성·이름 정의 보존, 원본과 백업 바이트 일치, 재검사 alreadyApplied를 확인했다.
손상 패키지는 주입 전 거부, 컴파일 오류는 저장하지 않고 원본 해시 유지, 첫 등록 계획은 파일 무변경이다.
테스트 변경: 새 GUI/프로세스 응답 경계의 추출 무결성·인자 전달·파일 변경 감지·저장/타임아웃 판정을
`UpdaterTests` 9개 사례로 추가했다. 기존 `PackageVerdictTests`에 오류 코드 단언 1개를 보강했으며 중복 사례는
추가하지 않았다. 단위 테스트의 Excel/프로세스 기동 증가는 0, 전체 974개 통과(442ms; 변경 전 시간은 미측정).
기존 publish 스모크를 GUI 1회 + 내장 CLI 2회 실행으로 대체해 정규 빌드당 프로세스 기동은 2회 늘었다.
coverage 기준은 변경하지 않았다. NuGet 취약성 목록은 검증 환경의 네트워크 제한으로 조회하지 못했다.

## 14. 실행 중단 (구현됨)

2026-11-03 벤치 시간 초과 보고에서 출발했다. 180초 제한에 걸리자 러너만 kill되고 매크로는 Excel 안에서 끝까지
돌았다(타임아웃 뒤에도 측정 기록 2건). 바쁜 Excel은 워크북 목록 판독(§11.2)에 실패해 읽기 전용 일회용 세션인데도
격리됐고, 사용자 확인과 재측정이 필요했다. 실측 근거는 `experiments/interrupt/README.md`.

### 14.1 문제
- 타임아웃이 작업을 멈추지 못한다. 결과는 "실패"인데 매크로는 완료됐을 수 있고, Excel은 바쁜 채 남는다.
- 타임아웃이 멈춤 감지(모달·행)와 예산을 겸한다. 호출자가 소요시간을 미리 추정해야 하고, 틀리면 비용이 크다.
- 명령 잠금(`CommandMutex`)은 부모만 쥔다. 부모가 강제 종료되면 잠금이 풀리지만 러너와 매크로는 계속 돌아,
  다음 명령이 같은 Excel에 끼어들 수 있다.

### 14.2 조사 결과
- 바깥에서 실행 중인 VBA를 멈추는 키 외 경로는 없다. VBE7.DLL(7.01.1143 x64)의 export는 `rtc*` 런타임과
  등록·초기화뿐이고 `EbMode`/`EbSetMode` 류가 없다. 바쁜 Excel은 COM 호출도 처리하지 않는다.
- VBE7은 `GetAsyncKeyState`를 import한다. 실행 중 Excel은 메시지를 읽지 않으므로(`IsHungAppWindow` true)
  중단 판정은 실제 키 상태 폴링으로 보인다. 게시형 `VK_CANCEL`(XLMAIN·EXCEL7·스레드)은 무시됐다.
  `SendInput`은 사용자의 전경 창에도 키가 가고, Ctrl+Break가 Excel 크래시를 낸다는 사용자 보고도 있어 쓰지 않는다.
- DLL 주입·디버거 조작은 버전 의존·보안 솔루션 차단·크래시 위험으로 제외한다.
- 협조적 취소는 동작한다. 확인 지점이 표시 파일을 보고 `End`하면 약 1초 안에 멈췄고, 러너가 정상 반환해
  격리 없이 일회용 Excel이 정리됐다. `Application.Run`은 `End`에 `0x800A9C68`(일반 런타임 오류와 같은 코드)을 낸다.
- COM 판독이 실패하는 바쁜 Excel에서도 `XLMAIN > EXCEL7` 자식 창 제목을 `InternalGetWindowText`(대상 스레드에
  메시지를 보내지 않음)로 47ms에 읽었다. 제목에는 `Probe.xlsm  [읽기 전용]`처럼 현지화된 접미사가 붙는다.

### 14.3 일회용 세션: 소유 증명 뒤 프로세스 정리 (구현됨)
- 일회용 인스턴스는 명령마다 버려지는 읽기 전용 Excel이라 종료해도 잃을 것이 없다. 지킬 대상은 그 인스턴스에
  다른 워크북이 열렸는지(DDE 재사용 등) 하나다.
- 취소·타임아웃 시 기존 신원 확인(PID + 시작 시각) 뒤 §11.2 프로브를 돈다. 프로브가 Busy/Unresponsive면
  EXCEL7 창 제목 판독으로 대신한다. 세션 워크북 창만 있으면 §11.1의 1번 규칙을 충족한 것으로 보고 kill한다.
  그 밖에는 지금처럼 격리한다.
- 창 제목 대조는 워크북 이름 기준이다. 현지화된 상태 접미사(`[읽기 전용]` 등)와 다중 창 접미사(`:1`)를 걷어 낸다.
- 실측(2026-11-03): 같은 인스턴스의 두 번째 워크북은 별도 XLMAIN 아래 EXCEL7(`통합 문서1`)로 보이고, 창을 숨긴
  워크북(`Windows(1).Visible = False`)도 EXCEL7이 남는다. 다중 창은 `Probe.xlsm:1  [읽기 전용]`, `:2`.
  코드가 Caption을 바꾸면 `Renamed`처럼 이름이 사라지므로 모르는 창으로 보고 격리한다. 숨은 인스턴스에서는 모든 창이
  보이지 않으므로 가시성은 근거가 아니다. 추가 기능(.xlam)은 창이 없어 보이지 않는다(사용자 작업이 아니다).
- 구현: `SessionProbe.Run`이 COM 목록을 못 얻으면(Busy/Unresponsive) 부모가 EXCEL7 제목을 읽어 `ProbeResult.WindowTitles`에
  붙인다(프로브 자식이 시간 초과로 죽어도 남는다). `RecoveryPolicy.DecideAuto`는 일회용이고 모든 제목이 세션 워크북
  창이면 Kill, 아니면 격리하고 이유에 창 제목을 적는다. 로그 `recovery_decision`에 `windowTitles`·`ephemeral`.
- 유지 세션은 이 근거를 쓰지 않는다. 세션 워크북 자체가 미저장 사람 작업을 담을 수 있다.

### 14.4 유지·attach 세션: LogShim 협조적 취소 (구현됨)
- Excel을 죽이면 사용자 상태까지 잃으므로 매크로가 스스로 멈추게 한다.
- 러너가 실행마다 고유한 표시 파일 `_dbg\cancel.<runId>`를 쓴다(이전 실행의 잔재가 다음 실행을 멈추지 않도록).
  러너는 시작 시 현재 runId를 LogShim에 알리고, 종료 시 파일을 지운다.
- LogShim의 기록 함수가 0.5초 간격(`Timer` 기준)으로 표시 파일을 확인한다. 있으면 `cancel` 기록을 남긴 뒤 `End`한다.
  러너는 이 기록으로 취소를 일반 런타임 오류와 구분해 `result: "cancelled"`로 보고한다.
- 기록을 남기지 않는 긴 루프용으로 확인만 하는 `LogShim.Poll`을 공개한다. 작성자가 루프에 직접 넣는다.
- 배포 빌드(`IS_RELEASE`)에서는 확인하지 않는다(§9).
- 정해진 유예 안에 멈추지 않으면 지금처럼 격리하고 사용자에게 넘긴다.
- 한계: 멈추기까지의 지연은 LogShim 호출 간격이다(보고서의 벤치는 약 20초). `End`는 모듈 수준 변수를 초기화한다.
- 반복문 머리에 확인 호출을 push 때 자동 삽입하는 방식은 보류한다. 측정값을 왜곡하고 src와 워크북 코드가 달라져
  pull/diff가 삽입분을 걸러야 한다. 위 두 경로로 부족한 사례가 나오면 검토한다.
- 구현: 러너가 실행(test는 테스트마다) 직전에 `'<wb>'!LogShim.BeginRun runId`를 부른다. 확인 지점은 `LogD`·`LogJ`
  (`LogEvent` 포함)·`AssertEq`·`Poll`이고, 취소 기록은 `{"kind":"X","event":"cancelled","runId":..}`다. 러너는
  호출 구간에 이 runId의 기록이 있으면 `cancelled`로 판정한다(run: `result: "cancelled"`, test: 그 테스트와 남은
  테스트는 skipped, 루트 `cancelled: true`).
- 구 LogShim과의 하위 호환은 두지 않는다. 러너는 실행 전에 워크북의 `LogShim` 모듈 본문에 `Public Sub BeginRun`이
  있는지 보고, 없으면(모듈 없음 포함) `vbaloop init` 후 push하라는 오류로 거절한다(exit 2). 취소할 수 없는 실행은
  Ctrl+C·`--timeout`이 곧바로 격리로 가므로 실행 자체를 막는 편이 낫다.
- `End`는 매크로의 정리 코드도 건너뛴다. 실측(2026-11-03): `Workbooks.Add` 뒤 취소된 매크로는 `wb.Close`에 닿지
  못했고, 일회용 정리가 남은 `통합 문서1`을 보고 Excel을 사용자에게 넘겼다(`ephemeral_handed_over`). §14.8이 이를
  다룬다.

### 14.5 취소를 일으키는 것 (구현됨)
실행할 때 넘기는 시간 예산을 기본으로 두지 않는다. `run`/`test`는 매크로가 끝날 때까지 기다린다.
- 모달은 타임아웃이 아니라 워치독이 즉시 감지해 실패로 끝낸다(VBE 오류 창 외의 MsgBox·InputBox·파일 대화상자·
  UserForm까지 확장).
- 호출자의 중단(Ctrl+C, 에이전트의 작업 중지)이 취소다. 부모가 받아 §14.3/§14.4로 정리한 뒤 끝낸다.
- 부모가 정리할 틈 없이 죽으면 러너가 부모 프로세스 핸들을 감시하다 같은 취소를 수행한다.
- `--timeout`은 스크립트·CI용 선택 옵션으로 남긴다. 기본값은 없다.
- 버려진 실행 정리용 상한은 프로젝트 설정에만 둔다(예: 1시간). 추정값이 아니라 아무도 취소하지 않는 경우의 안전망이다.
- 실행 중 진행 상황(경과 시간, LogShim 기록 수, 마지막 기록 이후 경과)을 주기적으로 출력한다. 호출자가 백그라운드로
  돌려도 진행·정체를 판단할 수 있게 한다.
- 구현(`RunnerSupervisor`): 250ms마다 러너 종료·취소 계기·대화상자를 본다. 취소 계기는 Ctrl+C(첫 번째), 선택
  `--timeout`, `[run] maxSeconds`(기본 3600) 중 먼저 오는 것이다. 계기가 오면 표시 파일을 쓰고 30초(`CancelGraceMs`)
  기다린다. 그 안에 러너가 돌아오면 정상 결과(cancelled)를, 아니면 강제 경로(러너 kill → poisoned → §11/§14.3)를
  탄다. 유예 중에 이 runId의 LogShim 취소 기록이 보이면(매크로는 멈췄다) 유예를 그 시점부터 다시 30초 준다 — 러너의
  뒷정리(§14.8, 계산 모드 복원의 재계산 등)가 성공한 취소를 강제 경로로 바꾸지 않게. 러너의 부모 감시도 같은 이유로
  취소 기록이 있으면 한 번 더 기다린다. 두 번째 Ctrl+C는 유예 없이 강제 경로다. 콘솔 Ctrl+C는 같은 콘솔의 러너에도
  가므로 러너도 `CancelKeyPress`를 막고 표시 파일을 쓴다.
- 막는 대화상자 판정(`ModalHunter.FindBlockingDialogs` + `BlockingDialogTracker`): 워크북 XLMAIN 중 하나라도
  비활성이면 그 프로세스의 최상위 `#32770`·`NUIDialog`·`ThunderDFrame` 중 VBE 오류 창이 아닌 것을 후보로 보고,
  같은 창이 5초 남으면 막힘이다. 실행 전부터 있던 창(attach한 사용자 Excel)은 뺀다. 실측(2026-11-03): 바쁜 루프는
  모든 XLMAIN이 활성, MsgBox 중에는 워크북 XLMAIN이 비활성이고 `#32770`(캡션 = MsgBox 제목)이 보인다. 그래서
  모드리스 UserForm은 걸리지 않는다. MsgBox 실행은 120초 대신 8초에 끝나고 오류에 캡션·버튼·본문이 실린다.
- 부모 사망: 부모가 러너에 `--parent-pid`를 넘긴다. 러너는 부모가 사라지면 표시 파일을 쓰고(`runner_parent_gone`)
  유예 동안 기다린 뒤에도 실행이 남으면 poisoned 기록과 §11 정리 후 종료 코드 3으로 끝난다(`runner_abandoned`).
  부모가 죽어 일회용 Excel을 닫지 못하면 다음 명령의 잔재 sweep이 정리한다.
- 진행 출력은 30초마다 stderr 한 줄이다(`vbaloop: run M.P still running - 1m00s elapsed, 60 LogShim record(s),
  last 0s ago`). 기록 수는 실행 시작 시점 이후 LogShim 파일 증가분이라 러너의 `run_start` 표식도 센다.

### 14.6 잠금 (구현됨)
- 러너가 실행 동안 프로젝트 실행 잠금을 쥔다(부모 사망과 무관하게 매크로가 끝날 때까지 유지).
- 다른 명령은 이 잠금을 보면 "실행 중인 run 있음"으로 거부하고 runId·경과 시간·진행 기록을 함께 보여 준다.
- 구현(`RunLock`): 커널 뮤텍스 `Local\vbaloop-run-<프로젝트 키>`와 기록 `_dbg\run-active.json`(runId·명령·대상·
  러너 PID+시작 시각·Excel PID·표시 파일·LogShim 기록 파일). 명령 잠금을 쥔 뒤 실행 잠금을 확인한다. 러너가 죽으면
  뮤텍스는 버려진 상태(= 비어 있음)가 된다. 러너 없이 남은 기록(부모·러너 트리 kill)은 다음 명령이 그 실행의 표시
  파일을 쓰고, 매크로가 멈췄다는 근거(이 runId의 취소 기록, Excel 소멸)가 나오거나
  30초(`OrphanWaitMs`)가 지날 때까지 exit 3으로 거부한다. 그 뒤 기록을 지우고 평소의 세션 복구에 맡긴다.
  격리 해제·session stop·격리 자동 해제는 남은 표시 파일을 모두 지운다(강제 경로 뒤 LogShim에 남은 runId가 사용자가
  계속 쓰는 Excel의 매크로를 End하지 않게).
- 러너의 부모 감시가 정리를 맡으면 본 스레드와 상태 플래그로 한쪽만 정리하고, 바쁜 Excel의 RCW를 쥐고 있으므로
  `ComCleanup.HardExit(3)`로 끝낸다. 러너의 Ctrl+C 처리기는 세션 확보 전에 건다(시작 중 Ctrl+C = 실행 전 취소).
- 막는 대화상자 판정은 LogShim 기록이 지속 시간 동안 늘지 않았을 때만 성립한다(Activate에서 일하는 모달 진행 폼). 부모는 강제 경로에서 러너를 kill한 뒤 자기 기록을 지우고, Excel이 정리됐으면
  표시 파일도 지운다. 하루 넘은 표시 파일은 러너가 시작할 때 지운다.

### 14.7 구현 순서
1. 일회용 세션 정리(§14.3): 완료. 시간 초과 시 바쁜 일회용 Excel이 격리 없이 정리된다.
2. 잠금(§14.6)과 부모 사망 감시(§14.5): 완료.
3. LogShim 협조적 취소(§14.4): 완료. 기존 프로젝트는 `vbaloop init` 후 push로 관리 사본을 갱신한다.
4. 모달 즉시 감지 확장과 기본 타임아웃 제거, 진행 출력(§14.5), 에이전트 문서 갱신: 완료.

E2E(2026-11-03, 일회용 세션, 60초 기록 루프 기준):
- `--timeout 10`: 약 1초 뒤 `cancelled`, 기록 보존, 격리 없음. `Poll`만 부르는 루프도 같다.
- LogShim 호출이 없는 루프: 유예 30초 뒤 강제 경로, 세션 창만 있어 Excel 정리(§14.3).
- MsgBox: 8초에 실패, 오류에 `E2E Box | 확인 | 취소 | e2e question`.
- Ctrl+C 1회: 기록 루프는 곧바로 `cancelled`, 호출 없는 루프는 유예 뒤 강제. 2회: 즉시 강제.
- 부모만 kill: 러너가 1초 안에 취소, 그 사이 다른 명령은 실행 잠금으로 거부(exit 3), 다음 명령은 정상.
- test `--timeout 8`: pass 2, 실행 중 테스트 skipped(cancelled while running), 남은 테스트 skipped.
- 두 번째 워크북이 열린 경우는 창 제목과 함께 격리(§14.3 회귀).

E2E(2026-11-04, 일회용 세션):
- `BeginRun`을 지운 LogShim: run·test 모두 실행 전에 `LogShim predates cancel support` + init/push 안내로 거절.
- test: 설정 5종을 바꾸고 새 문서를 연 뒤 런타임 오류로 끝난 테스트 → 다섯 설정 복원·문서 닫힘, 다음 테스트가 실행
  전 상태와 같음을 단언해 통과. 같은 일을 한 뒤 `--timeout`으로 취소된 테스트도 같다.
- run: 같은 매크로를 `--timeout 5`로 취소, 런타임 오류로 실패 → 둘 다 복원·문서 닫힘 뒤 일회용 Excel이 정상 종료.
  §14.4의 `Workbooks.Add` 뒤 취소 사례는 이제 넘기지 않고 닫는다.
- 유지 세션: 복원은 같고 문서는 닫지 않고 `left open`으로 보고한다. 정상 완료한 테스트가 남긴 문서는 뒷정리 대상이
  아니라 기존 규칙대로 넘긴다.

### 14.8 중단된 매크로 뒷정리 (구현됨)
취소와 런타임 오류(워치독이 [종료]로 해제)는 VBA `End`와 같다. 매크로의 정리 코드·`Class_Terminate`가 돌지 않아 세션
Excel에 상태가 남고, 같은 Excel을 쓰는 다음 테스트나 유지 세션의 다음 명령이 그 상태를 물려받는다.
- 러너가 매크로·테스트마다 호출 직전에 상태를 찍는다(`AppStateGuard.Take`): `EnableEvents`·`Calculation`·
  `StatusBar`·`Interactive`·`Cursor`와 열린 통합 문서 목록. `ScreenUpdating`·`DisplayAlerts`·`EnableCancelKey`는
  매크로가 끝나면 Excel이 되돌린다고 문서화돼 있어 뺀다(실측하지 않음).
- 결과가 취소·런타임 오류·모달 해제일 때만 되돌린다. 정상 완료한 매크로가 바꾼 설정은 의도로 보고 둔다. 세션이
  죽었으면 하지 않는다.
- 바뀐 설정은 실행 전 값으로 되돌린다. 이벤트는 끈 채로 되돌리고 `EnableEvents`는 마지막에 원래 값으로 둔다 — 계산 모드
  복원의 재계산이나 문서 닫기가 이벤트 코드(MsgBox 등)를 불러 뒷정리 호출을 막지 않게.
- 실행 뒤 새로 보이는 통합 문서는 일회용 세션이면 저장 없이 닫는다(세션 종료가 모르는 문서로 보고 Excel을 넘기지 않게).
  닫는 것은 잃을 것이 없는 문서뿐이다: 경로가 없는 새 문서, 또는 저장할 변경이 없는 문서. 경로가 있고 변경이 남은
  문서(매크로가 열어 고친 파일, 그 사이 사용자가 연 파일)와 SaveAs로 이름이 바뀐 세션 워크북(COM 개체 동일성으로
  가린다)은 닫지 않는다. 유지·attach 세션에서는 닫지 않고 보고만 한다.
- 결과는 run의 `payload.run.aftermath`, test의 각 테스트 `aftermath`(`restored`·`closed`·`leftOpen`·`failed`·
  `summary`)와 사람용 한 줄(`Cleanup after the interrupted macro: ...`, `NOTE <테스트> was interrupted: ...`)이다.
  뒷정리 실패는 판정을 바꾸지 않는다.
- 실측(2026-11-04): 한 번이라도 문구를 쓴 뒤 `StatusBar = False`로 돌려주면 Excel은 Boolean이 아니라 문자열
  `"FALSE"`를 돌려준다(VBA에서도 같다). 스냅숏은 이를 "Excel이 쥔 상태"로 본다.
- 다루지 않는 것: 반쯤 쓴 시트 데이터, 매크로가 닫은 기존 문서, `CreateObject`로 띄운 외부 앱, 임시 파일. 테스트는
  셋업을 멱등하게(있으면 지우고 다시 만들기) 짜서 정리 코드에 기대지 않는다.
