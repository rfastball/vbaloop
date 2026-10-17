# vbaloop dist - 배포 아티팩트 빌드

`vbaloop dist --out <경로> --json`은 배포용 클린 사본(`.xlsm` 또는 `.xlam`)을 만든다. on-disk
워크북을 base로 복사 -> 자체 임시 Excel(창 없음)에서 (모듈 제외 -> 버전 스탬프 -> LogShim 무해화
-> 선택 주석제거 -> VBE 컴파일 검증 -> 선택 부트스트랩)을 적용 -> 원자 이동으로 산출. 세션 워크북은
건드리지 않는다(lease 불필요). dist는 워크북을 바꾸지 않으므로 백업/rollback 대상이 아니다.

## 계약

- **push 선행 필수**: dist는 "마지막 push == src"인 on-disk 워크북을 base로 삼는다. 미동기면
  거부한다(`ok:false`, hint="vbaloop push로 반영 후 재시도"). **순서: (편집 ->) push -> dist.**
- 형식은 `--out` 확장자로 판정: `.xlsm`(매크로 워크북) / `.xlam`(애드인 - 자동 IsAddin + xlam
  저장). `--out` 생략 시 `<root>\dist\<워크북명>.xlsm`. `dist\`는 init-관리 .gitignore로 무시된다
  (빌드 아티팩트 = 비커밋).
- 제외 지정: 1차는 모듈 **헤더**의 `' @dist exclude`(본문에 적으면 무시된다 - 헤더 한정),
  보강은 toml `[dist] exclude` 글롭(`Test_*` 등). 제외는 모듈 단위만.
- 버전: toml `[dist] version`(semver). git rev는 자동 부가 - **git repo가 아니면 rev=`nogit`**
  (도구 자신의 rev를 스탬프하지 않는다). BUILD_* 상수 + CustomDocumentProperties 양쪽에 기록.
- 옵션: `--strip-comments`(주석 제거, 문자열 내 `'` 보호), `--bootstrap <Module.Proc>`(주입 후
  워크시트 초기설정 실행 - kill 워치독으로 모달 격리), `--skip-compile`(컴파일 검증 생략 - 탈출구).
- 컴파일 검증: 게시 전에 VBE "VBAProject 컴파일"을 실행한다. 결과는 `payload.compile.result`
  (`compiled` / `already_compiled` / `unverified`(경고 후 게시) / `compile_error` / `check_failed`).
  뒤의 둘이면 게시하지 않고(`payload.published: false`) 기존 배포본은 그대로 둔다.

## 실패 플레이북

- **sync 미동기(가장 흔함)**: hint대로 `vbaloop push` 후 재시도. **vbaloop 업데이트 직후 특히
  주의** - 새 exe의 `init`이 `src\LogShim.bas` 등을 재벤더링하면 src 지문이 어긋난다.
  **업데이트 -> init -> push -> dist** 순서를 지킨다(init은 순수 파일시스템이라 push 못 함).
- **dangling refs(컴파일 안전 실패)**: 제외 모듈의 심볼을 잔존 모듈이 참조. payload.danglingRefs의
  `모듈:라인`을 보고 제외 범위를 좁히거나 참조를 제거(Test 모듈은 서로만 참조해야 통째 제거 안전).
- **LogShim IS_RELEASE 부재(CliError)**: 배포 무해화 계약 위반 - 조용히 로깅 켜진 배포본을 막기
  위해 크게 실패한다. `src\LogShim.bas`가 관리 사본에서 벗어난 것이므로 `vbaloop init`으로
  정본을 되돌린다(그래도 재현되면 vbaloop 결함 - `vbaloop bugreport`).
- **컴파일 오류(`compile_error`)**: 에러 문구에 VBE 대화상자 본문이 담긴다(예: "변수가 정의되지
  않았습니다"). src에서 고치고 push -> dist. 제외 모듈의 심볼 참조가 원인이면 dangling refs와 같은 대응.
  test가 통과했는데 여기서만 걸리면 test가 닿지 않는 프로시저의 본문 오류다 - 그래도 배포본에선 터진다.
- **컴파일 검증 불가(`check_failed`)**: 컨트롤 없음/실행 실패/타임아웃. `doctor --json`으로
  AccessVBOM 확인. 검증 없이 게시해야 하면 `--skip-compile`(사람에게 알릴 것).
- **부트스트랩 타임아웃(120s)**: 부트스트랩 프로시저가 MsgBox/모달로 행을 건 것. 임시 Excel은
  kill로 회수된다(세션 무해). 해당 프로시저를 비대화형·컴파일 클린으로 고친다.
- **문서/시트 모듈 제외 불가**: warning으로 무시된다(Remove 불가). 진짜 제외가 필요하면 설계 재고.
- **`doctor --kill-orphans` 금지(dist 실행 중)**: dist의 임시 Excel은 창이 없어 doctor 고아 스캔에
  잡히고, doctor는 mutex를 안 잡으므로 dist를 죽일 수 있다. dist 진행 중엔 `--kill-orphans`를 쓰지 말 것.
