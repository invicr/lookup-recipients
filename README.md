# Excel Lookup Bundle

현재 폴더는 아래 3개 결과물만 기준으로 정리되어 있습니다.

- `dummy_1800.xlsx`: 1800명 더미 데이터
- `python/generate_lookup_html.py`: optimized HTML 생성 스크립트
- `html/lookup.html`: 조회용 HTML 결과물

기본 생성 명령:

```bash
python3 python/generate_lookup_html.py
```

기본 동작:

- 입력: `dummy_1800.xlsx`
- 출력: `html/lookup.html`

원하면 다른 파일로도 생성할 수 있습니다.

```bash
python3 python/generate_lookup_html.py /path/to/source.xlsx /path/to/output.html
```
