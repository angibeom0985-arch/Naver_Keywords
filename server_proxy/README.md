# Naver Keyword Proxy Server

클라이언트 EXE에서 네이버 API 키를 숨기기 위한 서버 중계용 API입니다.

## 1) 설치
```bash
cd server_proxy
pip install -r requirements.txt
```

## 2) 환경변수 설정
아래 값을 서버에 설정하세요.

- `SEARCHAD_ACCESS_KEY`
- `SEARCHAD_SECRET_KEY`
- `SEARCHAD_CUSTOMER_ID`
- `NAVER_CLIENT_ID`
- `NAVER_CLIENT_SECRET`
- `PROXY_TOKEN` (선택이지만 권장)

## 3) 실행
```bash
uvicorn main:app --host 0.0.0.0 --port 8080
```

## 4) 클라이언트 api_keys.json 예시
EXE 폴더의 `api_keys.json`은 아래처럼 설정하면 됩니다.
```json
{
  "proxy_url": "https://your-domain.com",
  "proxy_token": "your-strong-token",
  "usage_webhook_url": "",
  "usage_webhook_token": ""
}
```

`proxy_url`이 있으면 클라이언트는 네이버 API 키를 요구하지 않습니다.
