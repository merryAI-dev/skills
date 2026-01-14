# PPT Maker Skills

Claude Code 스킬 모음 - PPT 자동 생성 및 품질 관리 도구

## 스킬 목록

### 콘텐츠 생성

#### 1. ppt-generator
자연어 입력으로 PPT를 자동 생성합니다.

**사용법:**
```
"PPT 만들어줘: H-온드림 2026년 제안서, 총 15페이지"
```

**기능:**
- 자연어 → Claude API → slides_data.json → output.pptx
- 글자수 제한 자동 준수 (chapter: 5, title: 10-20, subtitle: 15-30, lead: 25-35)
- 문서 유형별 패턴 학습 (제안서, 결과보고서, 기획안, 백서)

#### 2. lead-inserter
기존 PPT에 리드문을 자동 생성하여 삽입합니다.

**사용법:**
```
"리드문 넣어줘: presentation.pptx"
```

**기능:**
- 슬라이드 구조 자동 추출 (chapter, title, subtitle)
- Claude API로 리드문 생성 (25-35자, 한 줄)
- 품질 평가 및 사용자 승인 워크플로우
- 원본 보존하며 새 파일 생성

### 품질 관리

#### 3. ppt-checker
3단계 파이프라인으로 맞춤법, 띄어쓰기, 용어 통일성을 검사합니다.

**사용법:**
```
"PPT 검사해줘: output.pptx"
```

**기능:**
- Stage 1: Ultra Think (XML + Claude API 구조 분류)
- Stage 2: Few-shot 모델 + 규칙 엔진
- Stage 3: Claude 문맥 검토 (False Positive 제거)
- Excel 리포트 생성 (변경 부분 빨간색 강조)

**특별 규칙:**
- "펠로우" → "펠로"
- "기업가 정신" → "임팩트 앙트프러너십"

### 템플릿 관리

#### 4. template-analyzer
템플릿 파일의 구조, 폰트, 색상, 위치 정보를 추출합니다.

**사용법:**
```
"템플릿 분석: tem.pptx"
```

**기능:**
- Placeholder 위치 및 크기 추출 (EMU, pt)
- 폰트 정보 추출 (이름, 크기, 색상)
- template_style.json 생성
- 커스텀 템플릿 제작 가이드

## 설치

이 스킬들은 PPT Maker 프로젝트와 함께 사용됩니다:

```bash
git clone https://github.com/merryAI-dev/pptMaker.git
cd pptMaker
pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."
```

## 통합 워크플로우

```bash
# 1. PPT 자동 생성
"PPT 만들어줘: 제안서 내용..."

# 2. 리드문 추가 (필요시)
"리드문 넣어줘: output.pptx"

# 3. 품질 검사
"PPT 검사해줘: output_with_leads.pptx"

# 4. 템플릿 커스터마이징 (필요시)
"템플릿 분석: custom_template.pptx"
```

## 기술 스택

- **Claude API**: claude-sonnet-4-20250514
- **python-pptx**: PPTX 파일 생성 및 조작
- **XML 파싱**: 구조화된 텍스트 추출
- **Knowledge Distillation**: Few-shot 예제 기반 품질 검사

## 라이선스

MIT License
