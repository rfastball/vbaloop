# vbaloop 에이전트 설치 절차

코드 에이전트(Claude Code, Codex 등)가 읽고 그대로 따르는 문서다. 사용자는 에이전트에게 이렇게 말하면 된다.

> https://github.com/rfastball/vbaloop/blob/master/AGENT_SETUP.md 를 따라 이 워크북에 vbaloop를 세팅해줘

## 원칙

- **동의 없이 바꾸지 않는다.** AccessVBOM(Excel 보안 설정)과 exe 설치·PATH 등록은 `setup --plan`이 `consent`로
  알려준다. 그대로 사용자에게 묻고 답에 따른다.
- **`--enable-vbom`은 사용자가 그 보안 설정 변경에 명시적으로 동의했을 때만 붙인다.** 레지스트리를 직접 고치지 않는다.
  `--yes`는 프롬프트만 생략하며 이 동의를 대신하지 않는다.
- **원본 워크북을 직접 편집·이동·삭제하지 않는다.** vbaloop가 원본 옆에 사본을 만든다.
- **Excel에 직접 손대지 않는다.** Excel 프로세스 종료, PowerShell `GetActiveObject`·`New-Object -ComObject` 같은
  COM 접근 모두 금지다. 환경 확인은 `vbaloop doctor --json`으로 한다.
- 모든 vbaloop 명령에 `--json`을 붙이고 봉투의 `ok`와 `hint`로만 분기한다. 실패하면 `hint`가 최우선이다.

## 1. 무엇을 세팅할지 정한다

| 상황 | 시작점 | 확인할 것 |
|---|---|---|
| 사용자가 워크북을 가리킴 | `--workbook "<원본.xlsm 절대 경로>"` | `.xlsm`인지. `.xlsx`/`.xls`면 사용자에게 Excel에서 `.xlsm`으로 다른 이름 저장을 요청 |
| 새로 시작 | `--new "<이름>"` | 이름, 그리고 만들 위치(이 명령을 실행하는 폴더 아래에 생긴다) |
| 준비해 둔 폴더 | `--folder "<폴더>"` | 그 폴더의 `.xlsm`이 0개(새로 만듦) 또는 1개(채택)인지 |

불분명하면 사용자에게 묻는다. 현재 폴더에 `.xlsm`이 하나뿐이면 그것을 후보로 제안해도 된다.

## 2. vbaloop 확보

```
vbaloop version
```

`vbaloop`를 찾지 못하면 `%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe version`을 시도한다. 둘 다 없으면 설치한다.
인터넷에서 설치 스크립트를 받아 실행하는 것이므로 사용자에게 알리고 진행한다.

```
powershell -NoProfile -Command "irm https://raw.githubusercontent.com/rfastball/vbaloop/master/install.ps1 | iex"
```

설치해도 **지금 에이전트가 쓰는 셸의 PATH는 바뀌지 않는다.** 이후 명령은 절대 경로
`%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe`(PowerShell: `& "$env:LOCALAPPDATA\vbaloop\bin\vbaloop.exe"`)로 실행한다.
아래의 `vbaloop`는 모두 이 경로로 읽는다.

## 3. 계획 보기 (아무것도 바꾸지 않는다)

1단계에서 정한 시작점으로 실행한다.

```
vbaloop setup --plan --json --workbook "C:\work\정산.xlsm"
```

`payload` 판독:

| 필드 | 의미 |
|---|---|
| `ready` | `false`면 `blockers`와 봉투의 `hint`를 사용자에게 전하고 멈춘다(Excel 미설치, 폴더 충돌 등) |
| `project` | 프로젝트 폴더(`folder`), 워크북, `action`(copy / reuse / inplace / create / adopt), 검증 방식(`verify`) |
| `changes` | 실행하면 바뀌는 것들. 사용자에게 보여줄 요약 |
| `consent[]` | 동의가 필요한 항목. `approveFlag`는 동의 시 붙일 플래그(null이면 이미 기본 동작), `declineFlag`는 거절 시 붙일 플래그(null이면 거절하면 진행 불가) |
| `command` | 동의 플래그를 뺀 실행 명령 |

## 4. 사용자 확인

`changes` 요약과 `consent` 항목을 한 번에 보여주고 묻는다. 예:

> 이렇게 세팅합니다.
> - `C:\work\정산\`에 원본 사본을 만들고(원본은 그대로) 코드를 `src\`로 꺼냅니다.
> - vbaloop를 `%LOCALAPPDATA%\vbaloop\bin`에 설치하고 PATH에 등록합니다.
> - **Excel 보안 설정 'VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스'를 켜야 합니다.** 꺼져 있으면 vbaloop가
>   동작하지 않습니다. 나중에 Excel 옵션 > 보안 센터 > 매크로 설정에서 끌 수 있습니다.
>
> 진행할까요?

- `accessVbom`을 거절하면 진행할 수 없다. 사용자가 Excel 옵션에서 직접 켠 뒤 다시 하자고 안내한다.
- `install`을 거절하면 `--no-install`을 붙인다. 이후에도 매번 절대 경로로 실행해야 한다고 알린다.

## 5. 실행

`command`에 동의한 `approveFlag`(예: `--enable-vbom`)와 거절한 `declineFlag`를 덧붙여 실행한다.
Excel을 띄워 검증하므로 몇 분 걸릴 수 있다. 타임아웃을 5분 이상으로 잡는다.

```
vbaloop setup --workbook "C:\work\정산.xlsm" --yes --json --enable-vbom
```

실패하면:

- `payload.consentRequired`에 `accessVbom`이 있음 → 4단계로 돌아가 동의를 받는다. 동의 없이 플래그를 붙이지 않는다.
- exit 3 → `hint`를 따른다. Excel이 셀 편집 중이거나 대화상자가 떠 있으면(busy) 사용자에게 알리고 재시도한다.
  격리(`recovery`) hint면 Excel을 죽이지 말고 사용자에게 상황을 알린다.
- 그 외 → `errors`와 `hint`를 사용자에게 그대로 전한다.

## 6. 마무리: `payload.next`

사용자에게 보고한다.

- 프로젝트 폴더(`next.folder`). 사본이면 원본은 그대로라는 점.
- 검증 결과(`verify`): `pull`이면 꺼낸 코드가 `src\`에 있다. `push+test`면 샘플 테스트가 통과했다.

이어가는 방법 두 가지를 제시한다.

- **새 세션(권장)**: `next.newSession`의 명령으로 프로젝트 폴더에서 에이전트를 연다. 그 폴더의 `CLAUDE.md`/`AGENTS.md`와
  스킬이 자동으로 로드된다.
- **이 세션에서 계속**: `next.continueHere`대로 관리 블록과 `SKILL.md`를 읽고 그 규칙을 따른다. vbaloop 명령은 프로젝트
  폴더에서 실행하거나 `--config`로 그 폴더의 `vbaloop.toml`을 지정한다. exe는 `next.exe`의 절대 경로를 쓴다.

`next.excelSession`이 있으면 검증용 Excel이 열려 있다. 바로 작업을 이어가지 않을 거면 프로젝트 폴더에서
`vbaloop session stop --json`으로 닫는다.

## 부록: 자주 만나는 상황

- **원본이 Excel에 열려 있음**: 마지막으로 저장된 내용이 복사된다. 저장하지 않은 변경이 있으면 사용자에게 먼저 저장하라고 권한다.
- **같은 워크북으로 다시 세팅**: 이전에 만든 개발 폴더를 재사용한다(`action: reuse`). 원본의 이후 변경은 반영되지 않는다.
  새로 가져오려면 사용자 확인 후 `--folder "<새 폴더>"`를 함께 지정한다.
- **폴더에 `.xlsm`이 여러 개**: `--workbook`으로 하나를 지정하면 그것만 새 폴더로 가져온다.
- **`--new` 폴더가 이미 있음**: 다른 이름을 쓰거나, 이어서 진행하려는 것이면 `--folder`로 그 폴더를 지정한다.
