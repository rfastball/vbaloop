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
  세션을 새로 띄웠을 때만 채워진다.
- 종료 코드: 성공 0 / 환경 오류 3(세션 부재·stale·poison·busy·격리, Excel/COM 치명) / 그 외 2.
  실패하면 한 줄 에러와 `hint`(다음 행동)를 준다.
- `_dbg\` 파일: `session.json`(세션 메타), `recovery.json`(격리 마커), `sync.json`(지문 캐시),
  `lease.json`, `vbaloop.jsonl`(append-only 도구 로그), `backups\`.
- 세션이 필요한 명령(push/pull/test/run)은 진입 시 세션을 스스로 확보한다(auto-ensure).

### doctor / status / diff
- `doctor`: Excel·COM·AccessVBOM·oleacc·비트니스·격리·고아 Excel·워크북 잠금 점검. `--fix`는 AccessVBOM을
  켜고, `--kill-orphans`는 §11.5 기준의 고아만 정리한다. 둘 다 명시할 때만 동작한다.
- `status`: 세션 상태(none/active/poisoned/stale/busy/quarantined), src와 마지막 push의 동기화 여부.
- `diff`: 워크북 vs src 읽기 전용 비교. 모듈별 same/changed/casingOnly/onlyInWorkbook/onlyInSrc와
  첫 차이 행을 보고한다. `--module`은 워크북 쪽 본문을 덤프한다. 차이가 있어도 exit 0이다(`payload.clean`).

### session
- `start`: 살아 있는 세션은 채택하고, 없으면 새 Excel을 띄운다. 이벤트·링크 갱신·보안 대화상자를 끈 채로
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
0. 지문 캐시: src가 마지막 push와 같으면 아무것도 하지 않는다(`--force`로 무시).
1. **백업**: 첫 변경 직전에 `SaveCopyAs`로 `_dbg\backups\<UTC id>-push\`에 사본과 `metadata.json`을 남긴다.
   `[backup] keep`(기본 10, 0은 비활성)개를 보존한다.
2. 파일 검사: BOM 제거, ACP 가드, 헤더 제거.
3. 모듈별 스냅샷 후 in-place 교체. 신규 모듈은 생성, 워크북에만 있는 모듈은 `--prune`일 때만 제거한다.
   문서 모듈(Sheet/ThisWorkbook)은 교체만 가능하다.
4. 무결성 검사: 예상하지 못한 컴포넌트 이름(`<이름>1` 등)이나 빈 모듈이 생기면 실패한다.
5. 저장. 저장하지 않은 성공은 보고하지 않는다.

실패하면 이번 push가 건드린 모듈을 전부 되돌리고 검증한다. 되돌릴 수 없으면(폼 재Import, prune,
무결성 실패, 원복 실패) 세션을 poison하고 저장하지 않는다. 디스크의 정본은 항상 push 이전 상태로 남는다.
`payload.recovery`에 결과를 담는다.

### rollback
`--list` / `--latest` / `--backup <id>`. 파일 작업만 한다. 세션 Excel이 살아 있으면 거부한다(exit 3).
현재 파일을 `pre-rollback` 백업으로 남긴 뒤 원자적으로 교체하고 해시를 대조한다. 이후 push는 전량 재주입하고,
pull은 `--force`를 요구한다.

### test / run
- `test`: 표준 모듈의 인자 없는 `Test_*` Sub를 모듈 한정으로 실행한다.
- 실행은 **자식 러너 프로세스**가 맡는다. `Application.Run`은 블로킹 호출이라 같은 프로세스에서는 시간 제한을
  걸 수 없기 때문이다. 부모가 벽시계 타임아웃(`--timeout`, 기본 120초)을 감시한다.
- 판정: 런타임 오류 없이 반환 AND 해당 구간의 LogShim `ok:false` 레코드 0건이면 pass.
- 어서션 싱크는 프로젝트에 벤더링한 self-contained `LogShim.bas`다. `Debug.Assert`는 헤드리스에서 아무 신호도
  내지 않는다.
- VBE 오류 대화상자는 러너의 워치독이 자동으로 닫는다(§8). 세션이 죽으면 남은 테스트는 skipped로 즉시 반환한다.
- `run <Module.Proc>`: 같은 러너로 단일 매크로를 실행하고 반환값을 보고한다.

### init
관리 아티팩트를 만들고 드리프트를 수렴한다. Excel이 필요 없고 멱등이다. `--check`는 검사만 한다.

| 아티팩트 | 관리 방식 |
|---|---|
| `vbaloop.toml` | 없으면 루트의 유일한 `*.xlsm`으로 스캐폴드. 사용자 파일이라 덮어쓰지 않는다 |
| `src\LogShim.bas` | 정본 사본(통째 관리) |
| `CLAUDE.md`, `AGENTS.md` | `<!-- vbaloop:managed:start/end -->` 블록만 관리 |
| `.claude\skills\vbaloop\` | 스킬과 참조 문서(통째 관리) |
| `.gitignore` | `# vbaloop:managed` 블록(`_dbg/`, `dist/`) |

정본은 exe에 임베드돼 있어 설치본만으로 온보딩할 수 있다. `--all`은 소스 레포 옆의 프로젝트들을 일괄 수렴한다.

### lint / bugreport / version / setup
- `lint`: 인코딩·규약·헤더·VB_Name·frm/frx 짝 전수 검사. Excel이 필요 없다.
- `bugreport`: 결함 리포트 템플릿을 만든다. 소스 레포에서는 `docs\bugreports\`, 설치본에서는 `_dbg\bugreports\`
  + GitHub 이슈 링크.
- `version`: semver + git 리비전 + 빌드 시각.
- `setup`: §10.

## 5. `vbaloop.toml`

```toml
[project]
name = "MyProject"            # 생략 시 워크북 이름

[excel]
workbook = "MyProject.xlsm"   # 루트 기준 상대 경로(필수)
visible = true

[src]
root = "src"

[dist]                        # 선택
version = "1.2.0"
exclude = ["Test_*"]
bootstrap = "modBootstrap.Bootstrap"
stripComments = false

[backup]                      # 선택
keep = 10
```

cwd에서 위로 올라가며 처음 만나는 `vbaloop.toml`의 위치가 프로젝트 루트다. 파서는 지원하는 부분집합
(문자열, bool, 정수, 문자열 배열, 주석)만 읽고, 그 밖의 문법은 조용히 무시하지 않고 오류로 거부한다.

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
7. test의 런타임·컴파일 오류 대화상자 처리와 타임아웃 격리

## 8. 부가 기능

- **모달 자동 회수**: 러너 안의 워치독이 VBE 대화상자를 찾아 닫는다. 버튼은 UI 언어와 무관하게 고른다.
  - 런타임 오류: 액세스 키 E(End). Debug는 절대 누르지 않는다(break 모드가 남는다).
  - 컴파일 오류: IDOK/IDCLOSE → 단일 버튼. 그 뒤 VBE를 닫아야 `Application.Run`이 반환된다.
  - 숨은 대화상자도 대상이다. 컴파일 직후에도 잠시 더 감시한다.
  - MsgBox 같은 사용자 대화상자는 닫지 않고 타임아웃으로 처리한다.
- **lease**: 변경 명령이 프로젝트 lease를 자동으로 잡는다. 소유자는 에이전트 호스트 프로세스(PID+시작 시각)
  또는 세션 id 환경변수다. 죽었거나 1시간 유휴인 소유자는 회수한다. 같은 프로젝트의 동시 명령은 OS mutex로
  직렬화한다.
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
- **9.3 LogShim 무해화**: 배포본에서만 `IS_RELEASE = True`로 바꾼다. 호출부는 그대로 두어 컴파일이 깨지지 않는다.
- **9.4 버전 스탬프**: 합성 모듈 `modBuildInfo`(상수)와 문서 속성(`vbaloop_BuildVersion`/`_BuildRev`/`_BuiltAt`).
- **9.5 주석 제거**: opt-in, 토크나이저 기반(문자열 안의 `'` 보존). 식별자 난독화는 하지 않는다(문자열로 결합된
  이름이 조용히 깨진다).
- **9.6 실행**: on-disk 워크북을 스테이징 파일로 복사해 별도 임시 Excel에서 변형하고, 성공하면 원자적으로
  게시한다. 실패하면 기존 배포본을 그대로 둔다.
- **9.7 형식**: `--out` 확장자로 정한다(`.xlsm` 또는 `.xlam`). 기본 위치는 `dist/`이고 init이 gitignore에 넣는다.
- **9.8 컴파일 검증**: 게시 전에 VBE 컴파일(CommandBars Id 578)을 실행한다. 컴파일 오류면 게시하지 않고
  대화상자 내용을 보고한다. `--skip-compile`로 건너뛸 수 있다.

## 10. 설치·배포

- **10.1** 대화형 명령은 사람용인 `setup` 하나뿐이다. 나머지는 에이전트용이라 비대화형이고 `--json`을 지원한다.
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
Excel을 종료하는 경로는 셋뿐이다.
1. 자동 정리(타임아웃·poison·stale): 도구가 띄운 세션 + PID·시작 시각 일치 + 세션 워크북 외에 열린
   워크북이 없음을 확인한 경우.
2. `session stop --discard`(명시): 도구가 띄운 세션 + 신원 일치 + 다른 워크북이 보이지 않는 경우.
3. `doctor --kill-orphans`(명시): §11.5의 고아.

그 밖의 경우는 모두 격리다.

### 11.2 프로브
응답하지 않는 Excel에 대한 COM 호출은 무기한 블록될 수 있다. 그래서 워크북 목록 확인은 자식 프로세스
(`__probe`)가 8초 제한으로 수행한다. 결과는 Responsive / Busy / Unresponsive / Windowless 중 하나다.

### 11.3 격리
kill할 수 없으면 `_dbg\recovery.json`을 남긴다. 이후 세션 명령은 exit 3과 해제 안내를 낸다.
- `recovery clear`: 기록된 프로세스가 사라졌을 때 해제한다(`--force`는 강제).
- `session stop --discard`: §11.1의 2번 규칙으로 정리하고 해제한다.
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
