# vbaloop

Excel VBA 프로젝트를 코드 에이전트가 다루게 해 주는 CLI.
src의 `.bas`/`.cls`/`.frm` 파일과 워크북 VBProject를 손상 없이 왕복시키고(push/pull),
`Test_*` 프로시저를 헤드리스로 실행해 결과를 JSON으로 돌려준다.
Windows 단일 exe이며 애드인이 필요 없다. 구현 근거는 [DESIGN.md](DESIGN.md).

## 특징

- **손상 없는 왕복**: 코드는 `CodeModule.Lines`(메모리)로 주고받아 코드페이지 변환이 없다. 한글 등 시스템
  ACP 문자는 그대로 보존되고, ACP로 표현할 수 없는 문자는 push가 저장 전에 거부한다.
- **안전한 push**: 주입 직전에 워크북을 백업하고, 실패하면 이번 push가 바꾼 모듈을 전부 되돌린다.
  `rollback`으로 백업에서 복원할 수 있다.
- **헤드리스 테스트**: 자식 프로세스에서 실행하고 타임아웃을 건다. VBE 오류 대화상자는 UI 언어와 무관하게
  자동으로 닫는다.
- **사용자 Excel 보호**: 사용자 작업이 있을 수 있는 Excel은 종료하지 않고 격리한 뒤 확인을 요청한다.
  이미 열어 둔 워크북은 `session attach`로 그대로 쓴다.
- **에이전트 연동**: `init`이 `CLAUDE.md`/`AGENTS.md` 관리 블록과 Claude Code 스킬을 설치하고 최신으로 유지한다.
  모든 명령이 `--json`을 지원한다.
- **배포본 빌드**: `dist`가 테스트 모듈 제외, 버전 스탬프, VBE 컴파일 검증을 거친 `.xlsm`/`.xlam`을 만든다.

## 요구 사항

- Windows + Excel 데스크톱(`.xlsm`)
- "VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스"(`setup`이 동의를 받고 켠다)
- 스킬은 Claude Code용이다. 다른 에이전트는 `AGENTS.md` 관리 블록을 쓴다

PC에 바꾸는 항목과 취약점 제보는 [SECURITY.md](SECURITY.md), 포함된 제3자 코드는
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)에 있다.

## 설치

```powershell
irm https://raw.githubusercontent.com/rfastball/vbaloop/master/install.ps1 | iex
```

최신 릴리스를 받아 SHA-256을 확인하고 `%LOCALAPPDATA%\vbaloop\bin`에 설치한 뒤 사용자 PATH에 등록한다.
또는 [릴리스](https://github.com/rfastball/vbaloop/releases/latest)에서 `vbaloop.exe`를 받아 더블클릭한다.
서명되지 않은 exe라 SmartScreen 경고가 뜨면 `추가 정보` → `실행`을 누른다.

## 시작하기

```powershell
vbaloop setup
```

환경 점검 → AccessVBOM → 설치 → 프로젝트 폴더 선택 → 워크북 생성 또는 채택 → `init` → 검증 순으로 진행한다.
신규 프로젝트는 샘플 테스트가 통과하는 것까지, 기존 워크북은 코드를 `src\`로 꺼내는 것까지 한다.
비대화형: `vbaloop setup --folder C:\work\MyVba --yes`.

그다음 그 폴더에서 코드 에이전트를 실행하면 된다. 에이전트는 관리 블록과 스킬을 읽고 편집 → push → test 루프를 돈다.

## 명령

| 명령 | 설명 |
|---|---|
| `setup` | 개발환경 마법사(유일한 대화형 명령) |
| `init [--check]` | 관리 파일 설치·갱신. `--check`는 검사만 |
| `doctor [--fix] [--kill-orphans]` | 환경 점검. `--fix`는 AccessVBOM 켜기, `--kill-orphans`는 창 없는 자동화 Excel 정리(먼저 `doctor`로 대상 확인) |
| `status` | 세션 상태와 src 동기화 여부 |
| `session start\|attach\|stop` | 세션 시작, 열린 워크북 채택, 종료(`--no-save`, `--discard`) |
| `push [--force] [--prune]` | src → 워크북 |
| `pull [--force]` | 워크북 → src |
| `diff [--module <이름>]` | 워크북과 src 비교(읽기 전용) |
| `test [--timeout <초>]` | `Test_*` 실행 |
| `run <Module.Proc>` | 매크로 1개 실행 |
| `rollback --list\|--latest\|--backup <id>` | push 전 백업으로 워크북 복원 |
| `recovery status\|clear [--force]` | 격리 상태 확인·해제 |
| `lint` | 인코딩·규약 검사(Excel 불필요) |
| `dist [--out <경로>]` | 배포본 빌드 |
| `lease acquire\|release\|status\|break` | 다중 에이전트 lease(보통 자동) |
| `bugreport "<제목>"` | 결함 리포트 작성 |
| `version` | 버전과 빌드 정보 |

- 모든 명령은 `--json`을 지원한다: `{command, ok, exitCode, hint, session, errors[], payload}`
- 종료 코드: 성공 0 / 환경 오류 3(세션·Excel·busy·격리, `hint`를 따른다) / 그 외 2

## 설정 (`vbaloop.toml`)

```toml
[project]
name = "MyProject"

[excel]
workbook = "MyProject.xlsm"
visible = true

[src]
root = "src"

[dist]                    # 선택
version = "1.0.0"
exclude = ["Test_*"]

[backup]                  # 선택
keep = 10                 # 0이면 push 전 백업 끔
```

`init`이 자동으로 만든다. cwd에서 위로 올라가며 처음 찾은 `vbaloop.toml`이 프로젝트 루트다.

## 테스트

- 표준 모듈의 인자 없는 `Public Sub Test_*()`가 테스트다.
- 단언은 `LogShim.AssertEq "이름", 기대값, 실제값`을 쓴다(`init`이 `src\LogShim.bas`를 넣는다).
  `Debug.Assert`는 헤드리스에서 아무 신호도 내지 않는다.
- 판정: 런타임 오류 없이 끝나고 실패한 단언이 0건이면 pass.

## 제약

- src 파일 규약: UTF-8(BOM 없음) + CRLF + 파일 끝 개행 1개.
- 프로시저 수준 Attribute(매크로 단축키 등)는 왕복되지 않는다. 단축키는 `Application.OnKey`로 등록한다.
- 폼의 `.frx`(그림 등)만 바뀐 변경은 감지되지 않을 수 있다. 이 경우 `push --force-form-import`로 반영한다.
- 테스트 중 `MsgBox`/`InputBox`/파일 대화상자는 자동으로 닫지 않는다. 타임아웃으로 끝나고, Excel은 격리될 수 있다.

## vbaflow에서 이전

`vbaflow` 시절 프로젝트도 그대로 동작한다. 각 프로젝트에서 `vbaloop init`을 한 번 실행하면 다음이 새 이름으로 옮겨진다.
- `vbaflow.toml`
- 관리 블록 마커
- `.claude\skills\vbaflow\`

dist가 기록하는 문서 속성 이름은 `vbaloop_*`로 바뀐다. 옛 설치 폴더 `%LOCALAPPDATA%\vbaflow\bin`은 직접 지운다.

## 개발

```powershell
dotnet build
dotnet test                                    # Excel 불필요
dotnet publish src\VbaLoop -c Release -r win-x64
```

.NET 10 SDK가 필요하다. 릴리스는 `v*` 태그를 push하면 GitHub Actions가 만든다.

## 라이선스

MIT
