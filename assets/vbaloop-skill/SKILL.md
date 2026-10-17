---
name: vbaloop
description: Excel VBA 프로젝트의 자율 개발 루프 절차. VBA 소스(.bas/.cls/.frm) 편집·주입·테스트·디버깅, .xlsm 워크북 매크로 작업, vbaloop 명령(session/push/pull/test/run/diff/lint/doctor/dist/rollback) 사용, 배포 아티팩트 빌드(dist), 실패 복구(poisoned/stale/busy/격리 세션, 사용자 Excel attach, lease 충돌, 타임아웃, 컴파일 오류 모달, push 부분 실패, 백업 복원)가 필요할 때 사용한다.
---

# vbaloop 워크플로우

vbaloop는 src 텍스트 파일과 Excel 워크북 VBA를 안전하게 왕복시키는 원자 명령 집합이다.
이 스킬은 원자 명령을 엮는 절차와 실패 플레이북을 담는다. 기본 규칙(인코딩, 관리 사본
수정 금지 등)은 CLAUDE.md 관리 블록에 있다. 명령 레퍼런스: `vbaloop help` (상세: https://github.com/rfastball/vbaloop).

## 증상 -> 조치 (인시던트 진입점)

곤경에 처하면 먼저 여기서 해당 절로 점프한다. 원칙은 항상 **응답 봉투의 `hint`가 최우선**이고,
아래 표는 hint를 해석할 맥락이다.

| 증상 | 조치 |
|------|------|
| 세션 없음/stale/오염 (exit 3) | auto-ensure가 정리+재기동. 반복 시 `doctor` -> §세션 poisoned/stale |
| 세션 격리 (exit 3, hint에 `recovery clear`/`--discard`) | 스스로 kill 금지, 사용자에게 알림 -> §세션 격리 |
| busy (exit 3, "셀 편집/대화상자") | 잠시 후 같은 명령 재시도 -> §세션 격리 |
| read-only 거부 + hint에 `session attach` | 사용자 확인 후 `session attach` -> §사용자 Excel attach |
| test: fail 1건 + skipped 다수 | 첫 fail의 `detail`(모달 본문)만 진짜 신호 -> §test 타임아웃/모달 |
| 고쳤는데 같은 오류, push가 계속 `cached:true` | `push --force`로 지문 캐시 우회 -> §push 거부 |
| push 거부 (CP949·모듈 불일치·문서모듈) | §push 거부 |
| push 중단 + `payload.recovery` | §push 실패 후 워크북 상태 |
| 워크북을 push 이전으로 되돌리고 싶다 | `session stop` 후 `rollback --latest` -> §백업과 rollback |
| pull 차단 (미동기 src 편집) | 기본 `push` 먼저, 워크북이 진실이면 `pull --force` -> §pull 주의 |
| lease 충돌 (다른 에이전트 보유) | 재시도/자동회수 대기 -> §lease 충돌 |
| dist 미동기·dangling·부트스트랩 타임아웃 | `references\dist.md` |
| 원인불명 컴파일/런타임 오류 | `references\vba-pitfalls.md` 조회 |
| 환경 의심 (Excel/COM/VBOM) | `doctor --json`이 최우선 |

## 판독 규칙 (모든 명령 공통)

- 항상 `--json`을 붙인다. 봉투: `{command, ok, exitCode, hint, errors[], payload}`
- 분기는 두 필드로만: `ok`(성공 여부)와 `hint`(다음 행동)
- exit code: 성공 0 / 환경 오류 3 (세션 부재·stale·오염, Excel/COM 치명, busy, 격리 - hint를 따른다) / 그 외 2
- 실패 시 `hint`가 다음 행동을 지시한다. hint를 따르는 것이 항상 최우선이다
- 상태를 기억하거나 추론하지 말 것 - "재검증 필요 여부"는 도구가 판단한다(해시 캐시, 멱등 start)

## 표준 루프

1. `vbaloop init --check --json` - 온보딩/드리프트 검사. 비정상(exit 2)이면 `vbaloop init`으로 수렴 후 진행
2. src\*.bas/.cls/.frm 편집
3. `vbaloop push --json` - 미변경이면 no-op. 실패 시 errors에 파일:라인 인라인(예: CP949 불가 문자)
4. `vbaloop test --json` - payload.tests[].detail에 실패 어서션(expected/actual)이 인라인됨
5. fail이면 detail로 원인 파악 후 2로. green(payload.summary.fail == 0 그리고
   payload.summary.skipped == 0)이면 git commit
6. `vbaloop session stop --json` - 저장 후 종료. 변경을 버릴 땐 `--no-save`

세션은 자동 확보된다(auto-ensure): push/pull/test/run이 세션 없음/stale/오염을 만나면
스스로 정리 후 기동하고, 새로 띄운 경우 봉투 `session: {action:"started", pid}`로 알린다.
`vbaloop session start`를 따로 부를 필요 없음(멱등이라 불러도 무해 - 진단/수동용).

테스트는 `Public Sub Test_*()`(인자 없음)만 수집된다. 단언은 `LogShim.AssertEq`.

## 실패 플레이북

### 세션 poisoned / stale

- 평시엔 만날 일이 줄었다 - 세션 필요 명령이 진입 시 스스로 정리+재기동한다(auto-ensure)
- exit code 3이면 자동 기동 자체가 실패한 것(Excel 기동/COM 치명) - JSON 파싱 없이도
  `session start` 후 재시도로 분기 가능. 반복되면 `vbaloop doctor --json`으로 환경 점검
- 치명적 COM 오류나 타임아웃 후 세션은 poisoned로 마킹된다. Excel이 이미 죽었거나 사용자 작업이
  없음이 확인되면 도구가 정리하고 다음 명령이 재기동한다 - 그렇지 않으면 격리된다(아래)
- 진단이 필요하면 `vbaloop status --json`으로 상태(none/active/poisoned/stale/busy/quarantined) 확인

### 세션 격리 / busy

- 도구는 Excel을 자동으로 kill하지 않는다. 타임아웃/poison 뒤 그 Excel에 사용자 작업이 없음을
  증명하지 못하면(다른 워크북이 열림, busy, 응답 없음, attach한 사용자 Excel) `_dbg\recovery.json`을
  남기고 이후 세션 명령은 exit 3 + hint로 거부된다
- **에이전트가 스스로 풀지 말 것.** `session stop --discard`(미저장 변경 유실)나
  `doctor --kill-orphans`로 Excel을 죽이지 말고, 사용자에게 상황(hint의 PID·사유)을 알린다:
  - 사용자가 그 Excel을 직접 저장/종료하면 -> `vbaloop recovery clear --json`
  - 사용자가 버려도 된다고 하면 -> `vbaloop session stop --discard --json`
    (도구가 띄운 Excel만 종료, 세션 외 워크북이 보이면 거부. attach 세션은 분리만)
- `vbaloop recovery status --json`으로 격리 사유와 프로세스 생존을 확인한다. 기록된 Excel이 이미
  종료됐으면 다음 세션 명령이 마커를 자동 해제한다
- `recovery clear --force`는 사람의 판단이 필요한 해제다 - 사용자가 지시할 때만
- busy(exit 3, "셀 편집 중이거나 대화상자가 열려 있음"): 사용자가 Excel을 쓰는 중이다.
  잠시 후 같은 명령을 재시도하고, 반복되면 사용자에게 알린다. 세션은 유지된다

### 사용자 Excel attach

- 사용자가 정본을 이미 자기 Excel에서 열어두면 auto-ensure는 read-only로 거부하고
  hint로 `vbaloop session attach`를 제시한다
- attach는 사용자 Excel 안에서 push 저장과 test/run 매크로 실행을 하게 되므로 **사용자 확인 후**
  `vbaloop session attach --json`. 이후 루프(push/test)는 평소와 같다
- `vbaloop session stop --json`은 분리만 한다(저장/닫기/종료 안 함 - push가 저장한 것만 남음)
- attach 세션의 타임아웃/poison은 항상 격리다 - 사용자 Excel은 어떤 경우에도 kill되지 않는다

### lease 충돌 (다중 에이전트)

- 변경성 명령(session/pull/push/test)이 자동 획득한다. 평시엔 신경 쓸 일 없음
- 다른 에이전트가 보유 중이면 즉시 exit 2 + 소유자/유휴시간이 표시된다
- 유휴가 짧으면 잠시 후 재시도. 소유자가 죽었거나 1시간 유휴면 자동 회수된다
- 확실히 버려진 lease만 `vbaloop lease break --force`로 강제 해제

### test 타임아웃 / 모달 / 세션 사망

- 컴파일 오류/런타임 오류 모달은 워치독이 자동 해제하고 해당 테스트만 fail 처리한다
  (detail에 모달 본문이 담김 - 이것이 실제 오류 메시지)
- 컴파일 오류 해제는 Excel 종료(세션 사망)를 유발할 수 있다. 이때 test는 즉시 중단하고
  잔여 테스트를 `result: "skipped"`로, `payload.sessionDied: true`로 보고한다
- **fail 1건 + skipped 다수 패턴이면 첫 fail의 detail(모달 본문)이 유일한 진짜 신호다** -
  skipped는 미실행일 뿐이니 원인 분석에 쓰지 말 것. 코드 수정 후 push + test로 재검증
  (세션은 자동 복구 안내를 따른다 - hint가 session start를 지시하면 그대로)
- MsgBox/무한루프는 타임아웃으로 끊긴다(기본 120초, `--timeout <초>`) - 러너만 종료되고 Excel은
  대개 격리된다(§세션 격리). 사용자에게 알리고 해제를 기다린다
- 테스트 코드에 MsgBox/InputBox를 넣지 말 것 - 헤드리스에서 행 유발

### push 거부

- CP949 불가 문자: errors의 파일:라인을 보고 해당 문자를 제거하거나 ChrW로 대체
- src에 없는 워크북 모듈: 의도한 삭제면 `--prune`, 아니면 pull로 먼저 회수
- 문서 모듈(ThisWorkbook/Sheet류)은 in-place 교체만 가능 - 신규 생성 불가
- 워크북과 src가 어긋난 의심(고쳤는데 같은 오류 반복, push가 계속 `cached:true`):
  `vbaloop push --force --json`으로 지문 캐시를 우회해 전량 재주입한다
  (`_dbg\sync.json` 수동 삭제는 불필요). 실물 확인은 `vbaloop diff --json`

### push 실패 후 워크북 상태

push는 실패하면 저장하지 않는다 - **디스크 정본은 항상 push 이전 상태**다. 메모리 워크북 상태는
`payload.recovery.workbookState`(그리고 에러 문구)가 말한다:

- `untouched`: 아무것도 안 바뀜. 원인 고치고 재시도
- `restored`: 이번 push가 건드린 모듈 전부 스냅샷으로 되돌리고 검증함. 원인 고치고 재시도
- `poisoned`: 되돌릴 수 없었음(폼 재Import·prune·무결성 실패·원복 실패). 세션은 버려지고 다음 명령이
  Excel을 **저장 없이** 재기동해 디스크 정본을 다시 연다 - 별도 조치 없이 원인 고치고 재시도
- 실패 직후 세션을 끝낼 땐 `session stop --no-save`가 안전하다(기본 stop은 저장한다)

### 백업과 rollback

- push는 주입 직전에 워크북 사본을 `_dbg\backups\<id>-push\`에 남긴다(no-op push는 생략,
  최신 10개 보존, toml `[backup] keep = N`, 0이면 비활성)
- 목록: `vbaloop rollback --list --json`. 복원: `vbaloop rollback --latest --json` 또는 `--backup <id>`
- **세션이 살아 있으면 exit 3** - `vbaloop session stop`(실패 직후면 `--no-save`) 먼저
- 복원 직전 현재 파일을 `pre-rollback` 백업으로 남긴다 - 바로 다음 `rollback --latest`가 되돌리기다
- 복원 후 src는 워크북과 다를 수 있다. **다음 push는 src를 전량 재주입해 복원을 덮어쓴다.**
  복원된 워크북을 src로 채택하려면 `vbaloop pull --force`(plain pull은 일부러 막혀 있다), 비교는 `vbaloop diff`

## pull 주의

- 동기화되지 않은 src 편집(지문이 마지막 push/pull 어느 쪽과도 불일치)이 있으면 pull이 차단된다.
  기본 대응은 `vbaloop push`로 먼저 반영. 워크북 쪽이 진실이면 `vbaloop pull --force`로 덮어쓴다
- 사람이 세션 Excel에서 직접 편집한 뒤에만 pull이 정당하다
- 확인만 필요하면 pull이 아니라 `vbaloop diff`(읽기 전용)를 쓴다 - pull은 src를 다시 쓴다
- 최초 pull(동기화 이력 없음)은 가드 없이 통과하며 일회성 포맷 정규화(BOM/개행) diff를 만든다 - 정상

## 폼 (.frm/.frx)

- 코드만 변경: 일반 push로 충분(in-place)
- 사람이 세션에서 폼을 수동 편집(특히 그림 교체): `vbaloop push --force-form-import --json`
- src에서 .frx를 직접 교체한 경우는 push가 자동 반영한다
- .frm 텍스트와 .frx는 짝 - lint가 짝 누락을 잡는다

## 왕복 한계 (설계상 불가 - 우회로 대응)

- 프로시저 수준 Attribute(VB_ProcData 매크로 단축키)는 왕복 불가, 최초 pull에서 소실.
  단축키는 Application.OnKey 런타임 등록으로 이전할 것
- 식별자 대소문자는 push 후 재pull에서 1회성 diff가 날 수 있다(수렴 후 안정 - 정상)
- 폼 그림은 frm 텍스트에 안 나타난다 - 수동 교체 감지 불가(위 --force-form-import로 대응)

## 배포 (dist)

`vbaloop dist --out <경로> --json`은 배포용 클린 사본(`.xlsm`/`.xlam`)을 만든다 - on-disk 워크북을
base로 복사해 임시 Excel(창 없음)에서 모듈제외·버전스탬프·LogShim 무해화·VBE 컴파일 검증을 적용 후 원자 이동.
**push 선행 필수**(미동기면 거부, hint="push로 반영 후 재시도"). 절차·형식(xlsm/xlam)·옵션·실패
플레이북 전체는 `references\dist.md`를 읽는다 - 편집 루프와 별개의 태스크라 상시 로드하지 않는다.

## VBA pitfall 참조 (조회/기록)

- **조회**: 신규 VBA 작성 전, 원인불명 컴파일/런타임 오류 디버깅 전에
  `references\vba-pitfalls.md`를 확인한다. 짧은 목록이라 훑는 비용이 싸다
- **기록**: 목록에 없는 pitfall로 디버깅 사이클을 태웠다면 상류에 제보한다. 프로젝트-로컬
  사본은 관리 파일이라 직접 수정해도 다음 init이 덮어쓴다:
  1. 기존 항목과 중복 확인 후, 문서 머리의 수록 기준·항목 포맷에 맞춰 초안을 만든다
  2. `vbaloop bugreport "<제목>" --json`으로 리포트를 남기고 초안을 본문에 채운다
     (설치본이면 payload.issueUrl의 GitHub 이슈로 제보 - 본문을 그대로 붙여넣으면 된다)
  3. vbaloop 소스 레포를 가진 개발 환경이라면 정본
     (`assets\vbaloop-skill\references\vba-pitfalls.md`)에 추가하고 다시 빌드한다
     (각 프로젝트는 다음 init --check가 드리프트를 감지해 수렴)

## vbaloop 자체 결함 리포트

- vbaloop 명령의 버그/설계 위반(불변식 깨짐, 잘못된 판정, 데이터 손상)을 만나면
  우회로 넘기지 말고 리포트를 남긴다: `vbaloop bugreport "<제목>" --json`
- payload.path의 생성 파일을 열어 TODO/괄호 자리(심각도/요약/증거/재현)를 채운다.
  증거는 실행한 명령과 --json 출력 원문을 코드블록으로. exe 빌드 정보는 자동 기입된다
- 중복 확인: `vbaloop bugreport --list --json` (resolved=false가 열린 리포트 - 같은 증상이면
  새 리포트 대신 기존 리포트에 증거를 추기)
- 설치본(소스 레포 없음)에서는 리포트가 `_dbg\bugreports\`에 생성되고 payload.issueUrl로
  GitHub 이슈 링크(https://github.com/rfastball/vbaloop/issues)를 함께 준다 - 본문을 채운 뒤 이슈로 제보한다

## 보조 명령

- `vbaloop doctor --json` - Excel/COM/AccessVBOM 환경 점검(세션 무관). 환경 의심 시 최우선.
  격리 상태와 창 없는 고아 EXCEL 프로세스도 리포트한다(정리 대상 여부와 사유 포함) - 정리는
  사용자가 확인한 뒤에만 `--kill-orphans`(에이전트가 먼저 실행하지 말 것)
- `vbaloop diff --json` - 워크북 VBProject vs src 읽기 전용 비교. 모듈별 verdict
  (same/changed/onlyInWorkbook/onlyInSrc/skipped)와 첫 차이 행(firstDiffLine)을 payload로 보고.
  워크북에 실제로 뭐가 들어있는지 확인할 때는 항상 이것 - 세션 Excel에 PowerShell 등으로
  직접 COM 접근은 금지. `--module <이름>`이면 해당 모듈의 워크북 쪽 본문 전체를
  payload.module.workbookBody로 덤프. 활성 세션 필요(attach 전용, 자동 기동 없음 - 없으면
  exit 3). 폼은 코드비하인드만 비교(레이아웃/.frx 제외). 차이가 있어도 exit 0 -
  payload.clean으로 판독
- `vbaloop lint --json` - 규약/인코딩/헤더 전수 검사(Excel 불필요). push 전 일괄 사전 점검
- `vbaloop run <Module.Proc> --json` - 임의 매크로 단건 실행. 반환값(원시 타입)/구간 로그 노출.
  테스트가 아닌 진입점 디버깅용 - 어서션 실패가 성패를 가르지 않는다
