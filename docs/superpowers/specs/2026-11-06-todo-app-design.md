# 할 일 목록 앱 설계

- 작성일: 2026-11-06
- 상태: 승인됨 (구현 계획 작성 전)

## 1. 목적과 범위

**목적:** 학습·포트폴리오용 할 일 목록 웹 앱. 기능보다 **코드 구조, 관심사 분리, 테스트**를 보여주는 것이 핵심이다.

**포함 기능**
- 할 일 추가, 제목 수정, 삭제
- 완료 여부 토글
- 필터: 전체 / 진행 중 / 완료
- 남은(미완료) 개수 표시
- 완료 항목 일괄 삭제
- 브라우저 `localStorage`에 저장 (새로고침 후에도 유지)

**범위 밖:** 배포, 마감일·우선순위, 정렬·드래그 순서 변경, 태그, 서버·계정·동기화

**전제:** 화면 문구는 한국어, 로컬 실행 기준, 모바일 화면 폭에서도 깨지지 않는 기본 반응형.

## 2. 기술 스택

- Vite + React 18 + TypeScript (`strict: true`)
- 상태 관리: `useReducer` + 순수 reducer + 커스텀 훅 (외부 상태 라이브러리 없음)
- 테스트: Vitest + React Testing Library + jsdom (+ `@testing-library/user-event`)
- 품질: ESLint, Prettier
- npm 스크립트: `dev`, `build`, `test`, `lint`

## 3. 데이터 모델

```ts
export type Todo = {
  id: string;        // crypto.randomUUID()
  title: string;     // 앞뒤 공백이 제거된, 비어 있지 않은 문자열
  completed: boolean;
  createdAt: number; // Date.now()
};

export type Filter = 'all' | 'active' | 'completed';

export type TodoAction =
  | { type: 'add'; id: string; title: string; createdAt: number }
  | { type: 'toggle'; id: string }
  | { type: 'edit'; id: string; title: string }
  | { type: 'remove'; id: string }
  | { type: 'clearCompleted' };
```

`id`와 `createdAt`은 reducer 밖(훅)에서 만들어 액션에 담는다. 그래야 reducer가 순수 함수로 남아 테스트에서 결과를 정확히 예측할 수 있다.

## 4. 구조

```
src/
  domain/
    types.ts            // Todo, Filter, TodoAction
    todoReducer.ts      // 순수 함수
    selectors.ts        // filterTodos, countActive
  storage/
    todoStorage.ts      // load(): Todo[], save(todos): void
  hooks/
    useTodos.ts         // useReducer + 저장 동기화 + 액션 함수
  components/
    TodoInput.tsx       // 새 할 일 입력
    TodoList.tsx        // 목록 렌더링
    TodoItem.tsx        // 체크박스, 더블클릭 수정, 삭제 버튼
    TodoFooter.tsx      // 남은 개수, 필터 버튼, 완료 항목 삭제
  App.tsx
  main.tsx
```

**의존 방향:** `components` → `hooks` → `domain`, `storage`. `domain`은 React에 의존하지 않는다. `storage`는 `domain/types`만 참조한다.

### 단위별 책임

| 단위 | 하는 일 | 공개 인터페이스 |
|---|---|---|
| `todoReducer` | 상태 전이 | `(state: Todo[], action: TodoAction) => Todo[]` |
| `selectors` | 파생 데이터 계산 | `filterTodos(todos, filter)`, `countActive(todos)` |
| `todoStorage` | 영속화와 데이터 검증 | `load()`, `save(todos)` |
| `useTodos` | 상태 보관, 저장 연동 | `{ todos, addTodo, toggleTodo, editTodo, removeTodo, clearCompleted }` |
| 컴포넌트 | 표시와 입력 | props로만 데이터와 콜백을 받음 (`App`만 훅 사용) |

## 5. 동작 흐름

1. 사용자 입력 → 컴포넌트가 `useTodos`의 액션 함수 호출 (예: `addTodo(title)`)
2. 훅이 `dispatch(action)` → `todoReducer`가 **새 배열**을 반환 (불변성 유지)
3. `useEffect`가 `todos`가 바뀔 때마다 `todoStorage.save(todos)` 호출
4. 화면 목록은 `filterTodos(todos, filter)` 결과. `filter`는 `App`의 `useState`이며 저장하지 않는다.
5. 처음 실행할 때 `useReducer`의 초기화 함수로 `todoStorage.load()`를 한 번 호출한다.

## 6. 규칙

- **추가:** 제목의 앞뒤 공백을 자른다. 빈 문자열이면 무시한다. 추가 후 입력창을 비운다. 새 항목은 목록 맨 아래에 붙는다.
- **수정:** 더블클릭으로 수정 모드에 들어간다. Enter나 포커스 이동(blur)은 저장, Esc는 취소. 앞뒤 공백을 자른 결과가 빈 문자열이면 **그 항목을 삭제**한다.
- **토글·삭제:** 해당 `id`가 없으면 상태를 그대로 둔다.
- **완료 항목 삭제:** 완료 항목이 하나도 없으면 버튼을 숨긴다.
- **남은 개수:** "남은 할 일 N개" 형식으로 표시한다.
- **빈 목록:** 할 일이 하나도 없으면 목록과 푸터 대신 안내 문구를 보여준다.

## 7. 예외 처리와 보안

- `load()`
  - 키가 없으면 → `[]`
  - JSON 파싱 실패, 배열이 아님 → `[]`
  - 배열 안에 형식이 맞지 않는 항목이 있으면 → 그 항목만 버린다 (타입 가드 `isTodo`)
  - `localStorage` 접근이 예외를 던지면 → `[]`
- `save()`: 예외(용량 초과, 사생활 보호 모드 등)를 `try/catch`로 잡아 무시한다. 앱은 메모리 상태로 계속 동작한다.
- 저장 키: `ssosso-todos-v1` (형식이 바뀌면 버전을 올린다).
- **XSS 방지:** 제목은 React의 기본 텍스트 렌더링으로만 출력한다. `dangerouslySetInnerHTML`은 쓰지 않는다.

## 8. 접근성

- 입력창과 체크박스에 레이블(`aria-label` 또는 `<label>`)을 붙인다.
- 필터는 `<button>`으로 만들고, 선택된 필터에 `aria-pressed="true"`를 단다.
- 마우스 없이 키보드만으로 추가·토글·삭제를 할 수 있어야 한다.

## 9. 테스트 전략

TDD로 진행한다: 실패하는 테스트를 먼저 작성하고, 통과시킨 뒤, 리팩터링한다.

| 대상 | 종류 | 확인할 내용 |
|---|---|---|
| `todoReducer` | 단위 | 액션별 결과, 없는 `id` 무시, 빈 제목 수정은 삭제, 원래 상태를 바꾸지 않음 |
| `selectors` | 단위 | 필터 3종 결과, 남은 개수 |
| `todoStorage` | 단위 | 저장 후 다시 읽기, 키 없음, 깨진 JSON, 잘못된 항목 거르기, `localStorage` 예외 |
| `App` | 통합 | 추가 → 완료 → 필터 → 완료 항목 삭제 흐름, 수정(Enter/Esc), 다시 마운트해도 데이터 유지 |

## 10. 완료 기준

- `npm test` 전부 통과
- `npm run build`, `npm run lint` 오류 없음
- 브라우저에서 기능 전체를 직접 확인
- README에 실행 방법, 폴더 구조, 설계 의도 설명
