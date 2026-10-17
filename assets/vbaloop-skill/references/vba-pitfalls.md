# VBA pitfall 참조 (에이전트용)

에이전트가 VBA를 작성/디버깅할 때 부딪히는 함정 목록. 신규 작성 전, 원인불명
컴파일/런타임 오류 앞에서 훑는다. 이 파일은 vbaloop 관리 사본 - 직접 수정 금지,
추가 제보는 `vbaloop bugreport` 또는 GitHub 이슈(https://github.com/rfastball/vbaloop/issues)로
(기록 절차: SKILL.md "VBA pitfall 참조" 절).

**수록 기준** (하나 이상 충족, 실제로 겪은 사례만):
(a) LLM이 타 언어 직관으로 그럴듯하게 틀리는 것
(b) 조용히/늦게 실패해서 디버깅 사이클을 태우는 것
(c) 헤드리스 자동화 세션 특유의 것
**항목 포맷**: `- **제목**: 증상/원인 1-2줄. 올바른 패턴.`
**제외**: 일반 스타일 조언, CLAUDE.md 관리 블록과 SKILL.md에 이미 있는 내용(링크만).

## 1. 헤드리스 세션 특유

- **이벤트 매크로가 실행되지 않는다**: 도구 세션은 `EnableEvents=False`로 워크북을
  연다 - `Workbook_Open`/`Worksheet_Change` 등에 의존하는 초기화는 조용히 생략된다.
  이벤트 의존 로직은 테스트에서 해당 프로시저를 명시 호출로 대신한다.
- **에러 탈출 시 Application 상태 잔류**: VBA엔 finally가 없다. `ScreenUpdating=False`,
  `Calculation=xlManual`, `EnableEvents=False`를 켠 채 에러로 빠지면 세션이 유지되는
  vbaloop 특성상 다음 테스트가 오염된 상태로 시작한다. 에러 핸들러에서 반드시 복원.
- **Selection/ActiveSheet 의존**: 자동화 세션에는 사용자가 만든 선택/활성 상태가 없다.
  비한정 `Range("A1")`은 ActiveSheet에 바인딩되어 다중 시트 코드에서 엉뚱한 시트를 쓴다.
  항상 `Worksheets("이름").Range(...)`로 한정.
- **Debug.Print는 헤드리스에서 안 보인다**: Immediate 창 출력은 판독 불가.
  터미널로 봐야 할 로그는 `LogShim.LogD`/`LogJ`. 단언은 `LogShim.AssertEq`.
- **중단점/Stop/VBE 디버그 모드 금지**: 수동 개입이 필요해 헤드리스에서 행이 된다.
  `Stop` 문을 소스에 남기지 말 것. 줄번호 없는 코드에서 `Erl`은 0을 반환한다.
- **파일 다이얼로그류도 행 유발**: MsgBox/InputBox(SKILL 참조)만이 아니라
  `Application.GetOpenFilename`/`GetSaveAsFilename`/`FileDialog`, `UserForm.Show`도
  헤드리스 실행 경로에서 금지 - 타임아웃으로만 끊기고 Excel은 대개 격리된다(사용자 해제 필요). 경로는 인자/상수로 주입.

## 2. VBA 언어 함정 (교차 언어 직관 배신)

- **And/Or는 단락 평가가 아니다**: `If x Is Nothing Or x.Count = 0 Then`은 x가 Nothing이어도
  `x.Count`를 평가해 런타임 91. 가드는 중첩 If로 분리: `If x Is Nothing Then Exit Sub`.
- **Dim i, j As Long에서 i는 Variant다**: 선언자마다 `As <타입>`이 필요하다.
  `Dim i As Long, j As Long`으로.
- **C-style 이스케이프는 없다**: `\"`, `\\`는 이스케이프가 아니라 리터럴 문자다.
  문자열 안 큰따옴표는 두 번: `"He said ""hi"""`. 경로 백슬래시는 그대로 `"C:\temp"`.
- **스마트 따옴표 혼입**: 생성/복붙 텍스트의 유니코드 따옴표는 VBA에서 구문 오류.
  ASCII 직선 따옴표만.
- **Integer는 16비트다**: 32,767에서 오버플로 - 행 카운트만으로 초과한다.
  정수는 기본 `Long`.
- **For Each 제어 변수는 Variant/객체 호환이어야 한다**: `Dim c As Long` 후
  `For Each c In coll`은 실패. `Dim c As Variant`(또는 요소의 객체 타입).
- **`Foo (arg)`의 괄호는 호출 문법이 아니다**: 인자를 값으로 강제 평가(ByRef 무력화)하거나
  객체의 기본 멤버를 평가한다. 프로시저 호출은 `Foo arg` 또는 `Call Foo(arg)`.
- **Me는 모듈 종류에 따라 없다**: Sheet/ThisWorkbook/클래스/폼 모듈에서만 유효, 표준
  .bas 모듈에는 없다. 어느 모듈 타입에 쓰는 코드인지 먼저 확인.
- **광역 On Error Resume Next는 오류를 삼킨다**: 실패가 조용히 전파되어 원인 특정이
  불가능해진다. 좁은 복구 블록 + 직후 `On Error GoTo 0`, 또는 라벨 핸들러
  (`Exit Sub`를 라벨 앞에). `Resume`은 핸들러 블록 안에서만.
- **Option Explicit 없으면 오타가 새 변수가 된다**: 모든 모듈 선두에 `Option Explicit`.
  미선언 대입은 컴파일 오류로 잡히는 것이 헤드리스에선 오히려 싸다(런타임 디버깅 회피).
- **라인 연속 `_` 앞에는 공백이 필요하다**: `arg1,_`는 구문 오류, `arg1, _`가 맞다.
- **End 문은 종류까지 짝이 맞아야 한다**: Sub를 `End Function`으로 닫으면 컴파일 오류 -
  헤드리스에선 컴파일 모달 경유라 비싸다. push 전 lint로 선검사.
- **AssertEq에 객체를 넘기지 말 것**: Range 등 객체는 기본 멤버 평가/91로 튀거나 의도와
  다른 비교가 된다. 스칼라로: `AssertEq "이름", 기대값, ws.Range("A1").Value2`.
- **모듈 수준 선언은 첫 프로시저보다 앞에만 올 수 있다**: 관련 프로시저 옆에 두려고
  프로시저 사이에 `Private m_x As Boolean`을 끼우면 컴파일 오류 "End Sub 다음에는
  주석만 나타날 수 있습니다". 변수/상수/Declare/Event 선언은 전부 모듈 선두 선언부에
  모으고 소속은 주석으로만 표기.

## 3. 다른 문서가 담당 (링크만)

- CP949 밖 문자, Debug.Assert 금지, 모듈 한정 호출, 테스트 green 전 커밋 금지:
  CLAUDE.md 관리 블록
- MsgBox/InputBox 타임아웃, 왕복 한계(VB_ProcData 소실, 케이스 1회성 diff, 폼 그림),
  poisoned/stale 복구: SKILL.md
