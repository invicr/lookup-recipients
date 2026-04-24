# Excel Lookup Bundle

현재 폴더는 아래 결과물을 기준으로 정리되어 있습니다.

- `dummy_1800.xlsx`: 1800명 더미 데이터
- `python/generate_lookup_html.py`: optimized HTML 생성 스크립트
- `html/lookup.html`: 조회용 HTML 결과물
- `html/lookup-data.json`: 조회용 암호화 데이터

기본 생성 명령:

```bash
python3 python/generate_lookup_html.py
```

기본 동작:

- 입력: `dummy_1800.xlsx`
- 출력: `html/lookup.html`, `html/lookup-data.json`

원하면 다른 파일로도 생성할 수 있습니다.

```bash
python3 python/generate_lookup_html.py /path/to/source.xlsx /path/to/output.html
```

주의:

- `lookup.html`은 `lookup-data.json`을 `fetch`로 불러오므로 `file://`로 직접 열지 말고 정적 서버에서 테스트해야 합니다.
