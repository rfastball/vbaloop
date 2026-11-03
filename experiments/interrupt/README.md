# 실행 중 매크로 중단·바쁜 Excel 판독 실험

2026-11-03. 시간 초과 보고(반납필요 검사 벤치 180초 초과, 러너만 kill되고 매크로는 Excel 안에서 계속 실행 →
워크북 목록 판독 실패 → 격리)의 근본 대책 후보 두 가지를 확인했다.

1. 실행 중인 VBA를 바깥에서 실제로 멈출 수 있는가
2. 매크로 실행으로 응답하지 않는 Excel에서도 COM 없이 열린 워크북을 판독할 수 있는가

## 환경과 방법

- 이 브랜치의 Release `vbaloop.exe`, `setup --new Probe --no-install --skip-verify --yes`로 만든 임시 프로젝트(스크래치 경로).
  Excel KR, 기본 ephemeral(숨은 읽기 전용 일회용 Excel).
- `modLoop.LongLoop`: DoEvents 없는 60초 바쁜 루프, 1초마다 `LogShim.LogJ "tick"`.
- `win32-probe.ps1`: Win32만 사용(COM 없음). `_dbg\ephemeral\*.json`의 PID로 대상 Excel을 특정하고,
  `XLMAIN > EXCEL7` 자식 창 제목을 `InternalGetWindowText`(대상 스레드에 메시지를 보내지 않음)로 읽은 뒤
  게시형 Ctrl+Break를 시도한다. 다른 Excel 프로세스는 건드리지 않는다.
- `modCancel.Check`: 협조적 취소. 0.5초 간격으로 `_dbg\cancel.flag`를 확인하고 있으면 기록 후 `End`.

## 결과

### 바쁜 Excel의 워크북 판독: 가능

| 항목 | 값 |
|---|---|
| 판독 소요 | 47ms (매크로 실행 중) |
| EXCEL7 제목 | `Probe.xlsm  [읽기 전용]` (현지화된 접미사 포함) |
| `IsHungAppWindow(XLMAIN)` | 실행 초반 false, 수 초 뒤 true |

COM(`AccessibleObjectFromWindow`)이 실패하는 바로 그 상황에서 창 계층으로는 판독된다.
`SessionProbe`의 폴백으로 쓰면 바쁜 일회용 Excel도 "세션 워크북만 열림"을 증명할 수 있다.

미확인: 두 번째 워크북이 같은 인스턴스에 열린 경우(DDE 재사용 등), 창을 숨긴 워크북(PERSONAL.XLSB 류),
창이 여러 개인 워크북(`Book:1` 제목), 추가 기능(.xlam은 창이 없어 이 방식으로 보이지 않음).

### 게시형 Ctrl+Break: 효과 없음

`PostMessage(XLMAIN)`, `PostMessage(EXCEL7)`, `PostThreadMessage(GUI 스레드)`로 `VK_CANCEL` keydown/keyup을
보냈지만 세 경우 모두 tick이 계속 늘었고 대화상자도 뜨지 않았다. 매크로는 60.0초에 정상 완료(`ok`).
VBA는 게시된 메시지가 아니라 실제 키 상태로 중단을 판정하는 것으로 보인다.
`SendInput`은 사용자의 전경 창(터미널 등)에도 Ctrl+Break가 가므로 제품 경로로 쓸 수 없어 시험하지 않았다.

### 협조적 취소(`End`): 동작

| 시각 | 사건 |
|---|---|
| 22:09:40.165 | `cancel.flag` 작성 (tick 5 직후) |
| 22:09:41.141 | 다음 tick의 `modCancel.Check`가 감지 → `cancel_seen` 기록 → `End` |
| 22:09:41.156 | 러너 `run_end result=fail` |

- `Application.Run`은 `End`에 대해 `0x800A9C68`을 반환했다. 일반 VBA 런타임 오류와 같은 코드라
  취소 여부는 `cancel_seen` 같은 직전 기록으로 구분해야 한다.
- 러너가 정상 반환했으므로 격리 없음, `recovery.json` 없음, 일회용 Excel 정리됨. elapsed 6017ms.
- 한계: 매크로가 확인 지점(LogShim 호출 등)을 지날 때만 멈춘다. 보고서의 벤치는 lap 기록 간격이 약 20초라
  지연이 그만큼 생기고, LogShim을 부르지 않는 순수 루프는 멈출 수 없다.

## 결론

- 실행 중 VBA를 바깥에서 강제로 멈추는 깨끗한 Win32 수단은 찾지 못했다.
- 협조적 취소(확인 지점 + `End`)는 확인 지점이 있는 매크로에서 확실히 동작하고 Excel을 정상 상태로 돌려놓는다.
- 확인 지점이 없거나 반응이 없으면, 바쁜 Excel에서도 COM 없이 워크북을 판독할 수 있으므로
  도구 소유 일회용 Excel이 세션 워크북만 열고 있음을 증명한 뒤 프로세스째 정리하는 폴백이 성립한다.

## 재현

```powershell
# 임시 프로젝트에서 modLoop.bas, modCancel.bas를 src에 넣고 push 후
vbaloop run modLoop.LongLoop --timeout 120 --json   # 백그라운드로
powershell -NoProfile -ExecutionPolicy Bypass -File experiments/interrupt/win32-probe.ps1 -Project <프로젝트 폴더>
# 협조적 취소: 실행 중 <프로젝트>\_dbg\cancel.flag 생성
```
