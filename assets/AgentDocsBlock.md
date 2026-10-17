## 개발 루프 (vbaloop)

빌드/테스트 루프는 `vbaloop`(단일 exe)로 수행한다. PATH 미등록 시 설치 경로:
`%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe` (그마저 없으면 `vbaloop setup`으로 재설치).
온보딩/드리프트 복구: `vbaloop init` (검사만: `vbaloop init --check`).
프로젝트 고유 내용은 이 관리 블록 밖에 작성한다 (init이 블록 내부를 덮어씀).

1. 편집: src\*.bas/.cls/.frm 수정
2. 세션: 자동 확보(push/test가 스스로 띄움, 수동은 vbaloop session start)
3. 주입: vbaloop push (src -> 워크북 in-place 주입, --prune/--force 지원. 주입 직전 _dbg\backups\에 백업)
4. 실행: vbaloop test (Test_* 전체 실행 + LogShim jsonl 판독, --timeout <초>)
5. 수정: fail이면 원인 고치고 1로, green이면 완료(버전관리를 쓰면 이때 커밋)

- 모든 명령 --json 지원: {command, ok, exitCode, hint, errors[], payload}
- exit code: 성공 0 / 환경 오류 3 (세션/Excel/COM/busy/격리 - hint를 따른다) / 그 외 2
- 격리(hint에 recovery clear / --discard): 도구가 Excel을 안전하게 정리할 수 없어 kill 대신 멈춘 상태.
  스스로 Excel을 죽이지 말 것(--discard, doctor --kill-orphans 포함) - 사용자에게 알리고, 사용자가
  그 Excel을 닫으면 vbaloop recovery clear, 버려도 된다고 하면 vbaloop session stop --discard
- busy(exit 3): Excel이 셀 편집/대화상자 중 - 잠시 후 같은 명령 재시도
- 사용자가 정본을 이미 Excel에서 열어둔 경우(read-only 거부 hint): 사용자 확인 후
  vbaloop session attach - 그 Excel에서 push/test가 동작하고 session stop은 분리만 한다
- 테스트는 Public Sub Test_*() (인자 없음), 단언은 LogShim.AssertEq (Debug.Assert 금지)
- src\LogShim.bas는 벤더링 사본 - 직접 수정 금지, 갱신은 vbaloop init
- push 실패 시 hint가 워크북 상태를 말한다(변경 없음 / 전부 원복 / 세션 poison - 저장 안 함).
  디스크 정본은 항상 push 이전 상태. 워크북 파일 자체를 되돌리려면 session stop 후
  vbaloop rollback --latest (목록: rollback --list, 세션이 살아 있으면 exit 3)
- 배포 아티팩트: vbaloop dist --out <경로> (클린 사본 빌드 - 필터/스탬프/무해화/VBE 컴파일 검증/선택 부트스트랩).
  push 선행 필수(마지막 push == src 미동기 시 거부). 절차/실패 플레이북은 SKILL.md
- pull은 동기화 안 된 src 편집이 있으면 차단된다 - 기본은 push로 먼저 반영,
  워크북 쪽이 진실일 때만 pull --force로 덮어쓰기
- 세션 Excel에 직접 COM 접근 금지(GetActiveObject, New-Object -ComObject 등 - 조회/진단 포함).
  워크북 실제 내용 확인은 vbaloop diff --json (읽기 전용, --module <이름>은 단일 모듈 덤프)
- src 규약: UTF-8 BOM 없음 + CRLF + 끝 개행 1개. CP949 밖 문자(em dash U+2014, 이모지)는 push가 거부
- vbaloop 자체 결함 발견 시: vbaloop bugreport "<제목>" 으로 표준 리포트 생성 후 내용 채움
  (절차: .claude\skills\vbaloop\SKILL.md)
- 심화 절차/실패 플레이북: .claude\skills\vbaloop\SKILL.md (vbaloop init이 관리 - 직접 수정 금지)
