# 경제 기사 요약 자동화 도구

매일경제 웹사이트에서 경제 기사를 크롤링하고, OpenAI GPT를 사용하여 기사를 요약한 후 Word 문서로 저장하는 자동화 도구입니다.

## 주요 기능

- 매일경제 경제 기사 자동 크롤링
- OpenAI GPT-3.5 Turbo를 활용한 기사 요약
- Word 문서로 자동 저장 (2열 표 형식)
- 기존 파일에 내용 추가 기능

## 설치 방법

### 1. 필요한 패키지 설치

```bash
pip install requests beautifulsoup4 openai python-docx
```

### 2. API 키 설정

`import requests.py` 파일에서 OpenAI API 키를 설정하세요:

```python
openai.api_key = 'your-api-key-here'
```

## 사용 방법

```bash
python "import requests.py"
```

## 출력 형식

각 기사마다 2열 표 형식으로 저장됩니다:

| 헤더 | 내용 |
|------|------|
| 날짜 | YYYY-MM-DD |
| 기사 제목 | [기사 제목] |
| 기사 링크 | [URL] |
| 기사 요약 | [GPT 요약 내용] |

## 파일 저장 위치

기본 저장 위치: `C:\Users\Desktop\경제신문.docx`

- 기존 파일이 있으면 내용을 추가합니다
- 기존 파일이 없으면 새로 생성합니다

## 주의사항

- 실행 전에 Word에서 `경제신문.docx` 파일을 닫아주세요
- OpenAI API 사용량에 따라 비용이 발생할 수 있습니다

## 라이선스

MIT License

<img width="732" height="661" alt="image" src="https://github.com/user-attachments/assets/26d44185-5a33-4492-839c-c4fffb52897b" />

