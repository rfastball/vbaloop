<h1 align="center">vbaloop</h1>

<p align="center">
  <b>Excel VBA를 코드 에이전트와 함께 개발하는 루프</b><br>
  편집 → push → test → 수정을 Claude Code·Codex가 사람 손 없이 돌게 해 주는 Windows CLI
</p>

<p align="center">
  <a href="https://github.com/rfastball/vbaloop/releases/latest"><img alt="release" src="https://img.shields.io/github/v/release/rfastball/vbaloop?include_prereleases&sort=semver"></a>
  <a href="https://github.com/rfastball/vbaloop/actions/workflows/ci.yml"><img alt="ci" src="https://github.com/rfastball/vbaloop/actions/workflows/ci.yml/badge.svg"></a>
  <img alt="platform" src="https://img.shields.io/badge/platform-Windows%20%2B%20Excel-0078D4">
  <a href="LICENSE"><img alt="license" src="https://img.shields.io/badge/license-MIT-green"></a>
</p>

<p align="center">
  <a href="#시작하기">시작하기</a> ·
  <a href="#작동-방식">작동 방식</a> ·
  <a href="#명령">명령</a> ·
  <a href="AGENT_SETUP.md">에이전트 설치 절차</a> ·
  <a href="DESIGN.md">설계</a>
</p>

> [!NOTE]
> **코드 에이전트에게 (For coding agents):** 사용자가 이 레포 주소와 함께 VBA·Excel 프로젝트의 설치나 세팅을
> 요청했다면 [AGENT_SETUP.md](AGENT_SETUP.md) 절차를 그대로 따른다. If a user gave you this repository to set up a
> VBA/Excel project, follow [AGENT_SETUP.md](AGENT_SETUP.md).

## 시작하기

### 준비물

- **Windows + Excel 데스크톱 버전.** Microsoft 365 웹·모바일 Excel은 자동화할 수 없다.
- **Claude Code 또는 Codex** (권장). vbaloop는 에이전트가 쓰는 도구다. 없어도 명령은 직접 쓸 수 있다.

새 프로젝트든 기존 워크북이든 시작 방법은 같다. 아래 둘 중 편한 쪽 하나만 하면 된다.

### 방법 1. 터미널에서 한 줄

**기존 워크북이 있는 폴더**, 또는 **새 프로젝트로 쓸 빈 폴더**에서 PowerShell을 열고 실행한다.

```powershell
irm https://raw.githubusercontent.com/rfastball/vbaloop/master/install.ps1 | iex
```

설치가 끝나면 그 자리에서 세팅이 이어진다. 폴더 상태를 보고 먼저 물어본다.

| 실행한 폴더 | vbaloop가 묻는 것 | 결과 |
|---|---|---|
| `정산.xlsm`이 있음 | "현재 폴더의 정산.xlsm을 가져올까요?" | 옆에 `정산\` 폴더를 만들어 **사본**으로 작업한다. 원본은 그대로다 |
| `.xlsm`이 여러 개 | 번호로 고르기 | 고른 워크북을 같은 방식으로 가져온다 |
| 비어 있음 | "여기에 새 프로젝트를 만들까요?" | 폴더 이름으로 워크북과 샘플 테스트를 만들고, 테스트가 통과하는 것까지 확인한다 |

Excel 보안 설정(AccessVBOM)이 꺼져 있으면 켤지 묻는다. 마지막으로 Claude Code가 설치돼 있으면 그 폴더에서 새 창으로
열지 묻는다.

### 방법 2. 에이전트에게 레포 주소 한 번

같은 폴더에서 Claude Code나 Codex를 열고 이렇게 말한다.

```text
https://github.com/rfastball/vbaloop 로 이 폴더에 vbaloop를 세팅해줘
```

에이전트가 [AGENT_SETUP.md](AGENT_SETUP.md)대로 진행한다.

1. vbaloop 설치
2. 현재 폴더를 보고 기존 워크북인지 새 프로젝트인지 판단
3. `setup --plan`으로 계획을 보여주고 동의 요청. Excel 보안 설정이나 PATH처럼 PC를 바꾸는 항목은 반드시 묻는다
4. 세팅 후 결과 보고

### 세팅이 끝나면

기존 워크북 `정산.xlsm`에서 시작했다면 폴더가 이렇게 된다.

```text
정산.xlsm                  원본 (건드리지 않음)
정산\                      ← 여기가 프로젝트. 에이전트를 이 폴더에서 연다
├─ 정산.xlsm               vbaloop가 다루는 사본
├─ src\                    워크북에서 꺼낸 모듈 (.bas .cls .frm) + LogShim.bas
├─ vbaloop.toml            프로젝트 설정
├─ CLAUDE.md, AGENTS.md    에이전트용 규칙 (vbaloop가 관리하는 블록)
├─ .claude\skills\vbaloop\ 실패 대응 절차 (Claude Code 스킬)
└─ .gitignore              _dbg\, dist\ 제외
```

이제 프로젝트 폴더에서 에이전트에게 평소처럼 일을 시키면 된다.

```text
modCalc의 합계 함수에 테스트를 붙이고 돌려줘
빈 시트에서 Run_Report 매크로가 오류 나는 거 고쳐줘
배포용 xlsm 만들어줘
```

에이전트는 `src\`를 고치고 `vbaloop push`로 워크북에 넣은 뒤 `vbaloop test`로 확인하고, 실패하면 다시 고친다.

<details>
<summary><b>다른 시작 방법</b> (exe 드래그 앤 드롭, 더블클릭, 명령 옵션)</summary>

<br>

이미 설치돼 있으면 워크북이 있는 폴더나 빈 폴더에서 `vbaloop setup`만 실행해도 방법 1과 같이 현재 폴더부터 제안한다.

- **기존 워크북**: `.xlsm`을 `vbaloop.exe` 위로 끌어다 놓는다.
- **새 프로젝트**: `vbaloop.exe`를 더블클릭하고 "기존 워크북을 가져올까요?"에 `n`, 이어서 이름만 입력한다.
  `%USERPROFILE%\vbaloop-projects\<이름>\`에 만든다.
- exe는 [릴리스](https://github.com/rfastball/vbaloop/releases/latest)에서 받는다. 서명되지 않은 exe라 SmartScreen 경고가
  뜨면 `추가 정보` → `실행`을 누른다.

명령으로 시작점을 직접 지정할 수도 있다.

```powershell
vbaloop setup --workbook C:\work\정산.xlsm    # 기존 워크북 → C:\work\정산\ 에 사본
vbaloop setup --new MyTool                     # 새 프로젝트 → .\MyTool\
vbaloop setup --folder C:\work\MyVba           # 준비해 둔 폴더(워크북 0~1개)를 그대로 사용
vbaloop setup --plan --workbook C:\work\정산.xlsm   # 아무것도 바꾸지 않고 계획만 보기
```

- 사본 폴더 이름이 이미 쓰이고 있으면 `<이름> (2)`처럼 비어 있는 이름을 쓴다. 다시 실행하면 이전에 만든 폴더를 재사용한다.
- 사본을 둘 곳을 직접 정하려면 `--workbook <경로> --folder <폴더>`.
- `--yes`는 묻지 않고 진행한다(Claude Code 실행 제안도 생략). 단 AccessVBOM을 켜는 것은 `--yes`에 포함되지 않는다.
  꺼져 있으면 `--enable-vbom`을 따로 붙여야 한다.
- 설치 한 줄에서 설치만 하고 세팅은 건너뛰려면 먼저 `$env:VBALOOP_NO_SETUP=1`을 설정한다. 에이전트·CI 셸(입력이 없는
  셸)에서는 원래 묻지 않고 에이전트용 안내만 출력한다.

</details>

## 작동 방식

```mermaid
flowchart LR
    A["src\ 편집<br/>.bas .cls .frm"] -->|vbaloop push| B[("워크북<br/>VBProject")]
    B -->|vbaloop test| C{"Test_* 결과<br/>JSON"}
    C -->|fail: 실패 내용 보고 수정| A
    C -->|green| D["완료 / 커밋"]
```

코드는 텍스트 파일(`src\`)이 원본이다. 에이전트가 파일을 고치면 vbaloop가 워크북 안의 VBA 프로젝트에 넣고, 테스트
프로시저를 헤드리스로 실행해 결과를 JSON으로 돌려준다. 에이전트는 그 결과만 보고 다음 행동을 정한다.

## 특징

| | |
|---|---|
| **손상 없는 왕복** | 코드는 `CodeModule.Lines`(메모리)로 주고받아 코드페이지 변환이 없다. 한글 등 시스템 ACP 문자는 그대로 보존되고, ACP로 표현할 수 없는 문자는 push가 저장 전에 거부한다 |
| **안전한 push** | 주입 직전에 워크북을 백업하고, 실패하면 이번 push가 바꾼 모듈을 전부 되돌린다. `rollback`으로 백업에서 복원할 수 있다 |
| **헤드리스 테스트** | 자식 프로세스에서 실행하고 타임아웃을 건다. VBE 오류 대화상자는 UI 언어와 무관하게 자동으로 닫는다 |
| **사용자 Excel 보호** | 사용자 작업이 있을 수 있는 Excel은 종료하지 않고 격리한 뒤 확인을 요청한다. 이미 열어 둔 워크북은 `session attach`로 그대로 쓴다 |
| **에이전트 연동** | `init`이 `CLAUDE.md`/`AGENTS.md` 관리 블록과 Claude Code 스킬을 설치하고 최신으로 유지한다. 모든 명령이 `--json`을 지원한다 |
| **배포본 빌드** | `dist`가 테스트 모듈 제외, 버전 스탬프, VBE 컴파일 검증을 거친 `.xlsm`/`.xlam`을 만든다 |
| **단일 exe** | 애드인이나 런타임 설치가 필요 없다. 설치 스크립트는 SHA-256을 확인한 뒤 `%LOCALAPPDATA%\vbaloop\bin` 한 곳에만 설치한다 |

## 테스트 작성

표준 모듈에 인자 없는 `Public Sub Test_*()`를 만들면 테스트가 된다. 단언은 `LogShim.AssertEq`를 쓴다.

```vb
Public Sub Test_SumTo()
    LogShim.AssertEq "1부터 3까지 합", 6, modCalc.SumTo(3)
    LogShim.AssertEq "0이면 0", 0, modCalc.SumTo(0)
End Sub
```

- 판정: 런타임 오류 없이 끝나고 실패한 단언이 0건이면 pass.
- `src\LogShim.bas`는 `init`이 넣고 관리한다. 직접 고치지 않는다.
- `Debug.Assert`는 헤드리스에서 아무 신호도 내지 않으므로 쓰지 않는다. 테스트 안에서 `MsgBox`·`InputBox`도 쓰지 않는다.

## 명령

| 명령 | 설명 |
|---|---|
| `setup [--workbook <xlsm> \| --new <이름> \| --folder <경로>] [--plan] [--enable-vbom]` | 개발환경 마법사(유일한 대화형 명령). `--plan`은 미리보기 |
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

`setup`(`init`)이 자동으로 만든다. cwd에서 위로 올라가며 처음 찾은 `vbaloop.toml`이 프로젝트 루트다.

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

## 제약

- src 파일 규약: UTF-8(BOM 없음) + CRLF + 파일 끝 개행 1개.
- 프로시저 수준 Attribute(매크로 단축키 등)는 왕복되지 않는다. 단축키는 `Application.OnKey`로 등록한다.
- 폼의 `.frx`(그림 등)만 바뀐 변경은 감지되지 않을 수 있다. 이 경우 `push --force-form-import`로 반영한다.
- 테스트 중 `MsgBox`/`InputBox`/파일 대화상자는 자동으로 닫지 않는다. 타임아웃으로 끝나고, Excel은 격리될 수 있다.

## 보안

vbaloop가 PC에 바꾸는 것은 셋이다: Excel의 "VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스"(AccessVBOM),
`%LOCALAPPDATA%\vbaloop\bin` 설치와 사용자 PATH 등록, 프로젝트 폴더의 관리 파일. AccessVBOM은 마법사에서 동의하거나
`--enable-vbom`을 명시해야만 바뀌고, 무엇이 바뀌는지는 `setup --plan`으로 미리 볼 수 있다. 원본 워크북은 수정하지 않는다.
자세한 내용과 되돌리는 방법, 취약점 제보는 [SECURITY.md](SECURITY.md),
포함된 제3자 코드는 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)에 있다.

<details>
<summary><b>vbaflow에서 이전</b></summary>

<br>

`vbaflow` 시절 프로젝트도 그대로 동작한다. 각 프로젝트에서 `vbaloop init`을 한 번 실행하면 다음이 새 이름으로 옮겨진다.

- `vbaflow.toml`
- 관리 블록 마커
- `.claude\skills\vbaflow\`

dist가 기록하는 문서 속성 이름은 `vbaloop_*`로 바뀐다. 옛 설치 폴더 `%LOCALAPPDATA%\vbaflow\bin`은 직접 지운다.

</details>

## 개발

```powershell
dotnet build
dotnet test                                    # Excel 불필요
dotnet publish src\VbaLoop -c Release -r win-x64
```

.NET 10 SDK가 필요하다. 릴리스는 `v*` 태그를 push하면 GitHub Actions가 만든다. 구현 근거는 [DESIGN.md](DESIGN.md).

## 라이선스

[MIT](LICENSE)
