# GPT for Sheets (Apps Script)

Google Sheets에서 직접 OpenAI GPT 모델을 호출해 프롬프트와 시트 데이터를 분석·요약할 수 있도록 돕는 Apps Script 프로젝트입니다. `Code.gs`, `Config.gs`, `Utils.gs` 파일을 이용해 커스텀 함수 `GPT`와 `GPT_JSON`을 제공하며, 캐시와 오류 처리까지 포함한 일체형 솔루션입니다.

## 사전 준비
- OpenAI API 키 (예: `sk-` 로 시작)
- Google 계정 및 Google Sheets 접근 권한
- Google Apps Script 편집 권한 (시트에 바인딩된 스크립트 또는 독립 Apps Script 프로젝트)

## 파일 구성
- `Code.gs` : 커스텀 함수 본체 및 OpenAI API 호출 로직
- `Config.gs` : API 키 저장, 캐시 버전 관리, 설정용 유틸 함수
- `Utils.gs` : 범위 직렬화, 캐싱, 오류 정규화 등 공용 유틸리티

## 설치 및 설정
1. Google Sheets에서 `확장 프로그램 > Apps Script` 를 열고, 기존 코드를 모두 삭제한 뒤 위 세 파일의 내용을 복사해 붙여넣습니다.
2. 최초 한 번, Apps Script 에디터에서 `setApiKey('sk-...')` 를 실행하여 OpenAI API 키를 저장합니다.  
   - 실행 시 권한 요청이 나타나면 승인해야 합니다.  
   - UI로 입력하고 싶다면 `configureApiKey()` 를 실행해도 됩니다.
3. (선택) 캐시 초기화를 원하면 `flushCache()` 를 실행합니다.
4. 시트로 돌아가 `=GPT("테스트")` 와 같이 함수를 호출해 정상 동작을 확인합니다.

## 시트 함수 사용법

### `GPT(prompt, [rangeInput], [options])`
OpenAI Chat Completions API를 호출해 문자열 응답을 반환합니다.

- `prompt` (필수, 문자열) : 모델에 전달할 지시문.
- `rangeInput` (선택, 시트 범위 또는 2차원 배열) : 추가 컨텍스트로 전달할 데이터.
- `options` (선택, JSON 문자열 또는 객체) : 모델/출력 제어 옵션.

#### 지원 옵션
| 옵션 키 (camelCase / snake_case) | 기본값 | 설명 |
| --- | --- | --- |
| `model` | `gpt-4o-mini` | 사용할 모델 ID |
| `temperature` | `0.7` | 생성 다양성 (0~2) |
| `topP` / `top_p` | `null` | 확률 질량 상위 비율 (0~1) |
| `maxTokens` / `max_tokens` | `null` | 응답 토큰 수 상한 (최대 16000) |
| `format`, `response_format` | `text` | `json` 지정 시 JSON 응답 강제 |
| `responseCharLimit`, `response_char_limit` | `2000` | 셀에 출력될 최대 문자 수 |
| `cacheTtlSeconds`, `cache_ttl_seconds` | `3600` | 캐시 TTL(초), 상한 21600 |
| `systemPrompt`, `system_prompt` | 기본 시스템 프롬프트 | 시스템 메시지 교체 |
| `tokenLimit`, `token_limit` | `16000` | 프롬프트+응답 추정 토큰 상한 |
| `frequencyPenalty`, `frequency_penalty` | `null` | 반복 억제 (-2~2) |
| `presencePenalty`, `presence_penalty` | `null` | 신규 토픽 장려 (-2~2) |

#### 사용 예시
- 단순 호출: `=GPT("이 데이터의 인사이트를 요약해줘")`
- 범위 전달: `=GPT("다음 표를 요약해줘", A1:C10)`
- 옵션 사용: `=GPT("영문으로 3줄 요약", A1:C10, "{\"temperature\":0.2,\"maxTokens\":200}")`
- 캐시 무시(응답 길이 확대): `=GPT("긴 보고서를 작성해줘", A1:C50, "{\"responseCharLimit\":4000,\"cacheTtlSeconds\":10}")`

### `GPT_JSON(prompt, [rangeInput])`
`GPT` 함수의 래퍼로, `{"format":"json"}` 옵션이 자동 적용됩니다. JSON 형태의 응답이 필요할 때 사용하며, 시트에서는 `=PARSE_JSON()` 등의 함수와 함께 사용할 수 있습니다.

## 관리용 함수
- `setApiKey(key)` : OpenAI API 키 저장. (Apps Script 에디터에서 실행)
- `configureApiKey()` : 시트 UI 팝업으로 키 입력.
- `getApiKey_()` : 내부용 API 키 조회.
- `flushCache()` : 캐시 버전을 갱신해 모든 캐시 무효화.
- `clearApiKey()` : 저장된 키를 삭제.

## 캐시 및 토큰 정책
- 캐시 키는 프롬프트, 직렬화된 범위 데이터, 시트 이름, 주요 옵션, 캐시 버전으로 구성되어 동일 호출의 재사용을 극대화합니다.
- 기본 캐시 TTL은 1시간이며 옵션으로 조정 가능합니다.
- 토큰 추정은 문자 길이 기반 휴리스틱을 사용하며, `tokenLimit` 을 초과하는 경우 `#GPT_ERROR: TOKEN_LIMIT` 오류로 호출을 중단합니다.

## 오류 코드 안내
| 코드 | 의미 |
| --- | --- |
| `#GPT_ERROR: Missing prompt` | 프롬프트가 비어 있음 |
| `#GPT_ERROR: API_KEY_MISSING` | 저장된 OpenAI API 키가 없음 |
| `#GPT_ERROR: TOKEN_LIMIT - estimated ...` | 예상 토큰 수가 허용치를 초과 |
| `#GPT_ERROR: REQUEST_FAILED` | HTTP 오류 또는 응답 파싱 실패 |
| `#GPT_ERROR: No response` | API가 비어 있는 메시지를 반환 |
| 그 외 `...(detail)` | 구체적 오류 메시지가 함께 표시됨 |

## 문제 해결 팁
- **권한 오류** : Apps Script에서 최초 실행 시 나오는 권한 요청을 반드시 승인해야 합니다.
- **API 키 관련 오류** : `clearApiKey()` 로 초기화 후 다시 `setApiKey()` 를 실행합니다.
- **캐시된 응답 변경 필요** : `flushCache()` 실행 또는 옵션의 `cacheTtlSeconds` 를 낮춥니다.
- **JSON 파싱 실패** : `GPT_JSON` 사용 시 모델이 유효한 JSON을 생성하도록 프롬프트에 명시하고, 응답이 JSON 형식을 유지하도록 재시도합니다.

## 추가 참고
- 응답이 너무 길면 `responseCharLimit` 내에서 잘리고 `...(truncated)` 가 붙습니다.
- 범위 직렬화 시 빈 행은 자동으로 제거되며, 날짜/객체는 ISO 포맷 또는 JSON 문자열로 변환됩니다.
- 모든 함수는 스크립트 속성에 API 키를 저장하며, 코드에 키를 직접 하드코딩하지 않습니다.


