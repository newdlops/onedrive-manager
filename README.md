# OneDrive 관리자

Windows와 macOS에서 동작하는 Microsoft OneDrive 원격 파일 관리자입니다.

## 제품 방향

이 앱은 OneDrive 동기화 클라이언트가 아닙니다. OneDrive에 있는 파일을 원격 저장소처럼 탐색하고, 필요한 파일을 직접 내려받거나 백업할 파일을 직접 올리는 수동 파일 관리 도구입니다.

자동 동기화, 실시간 폴더 감시, 로컬 OneDrive 동기화 상태 제어는 제품 범위에서 제외합니다.

## 1차 목표

Microsoft 계정으로 OneDrive에 접속하고, 계정의 OneDrive 파일과 폴더를 원격 탐색기 형식으로 확인하는 것을 1차 목표로 둡니다.

상세 범위는 [docs/phase-1.md](docs/phase-1.md)에 정리합니다.

## 언어 정책

현재 앱은 한국어만 지원합니다. 사용자에게 보이는 문구는 한국어로 작성하고, 다국어 지원이 실제 요구사항이 되기 전까지 i18n 추상화는 추가하지 않습니다.

## 플랫폼 정책

- OS별 동작은 Electron main process의 작은 모듈 안에 둡니다.
- 네이티브 의존성을 추가하기 전에 표준 Node/Electron API를 우선 사용합니다.
- macOS에서 주로 개발하더라도 Windows를 주요 실행 환경으로 취급합니다.
- Windows와 macOS 네이티브 CI 러너에서 빌드와 패키징을 검증합니다.
- renderer에서 파일시스템에 직접 접근하지 않고 preload bridge와 타입이 지정된 IPC를 사용합니다.

## 스크립트

```sh
npm run dev
npm run typecheck
npm run build
npm run package
npm run dist:mac
npm run dist:win
```

`dist:win`은 macOS에서도 기본 cross-build 용도로 실행할 수 있지만, 설치 프로그램과 서명, 파일시스템 동작은 플랫폼 영향을 받으므로 릴리스 빌드는 Windows CI에서 확인합니다.

`dist:mac`은 Intel Mac과 Apple Silicon Mac을 모두 지원하는 universal DMG를 `release/`에 생성합니다.

## Microsoft 계정 로그인 설정

Microsoft Entra에서 데스크톱 public client 앱을 등록한 뒤 빌드 환경에 Application client ID를 주입합니다. Application client ID는 OAuth public client 식별자이며 client secret이 아니지만, 저장소에는 올리지 않습니다.

개발 중에는 로컬 `.env`를 사용할 수 있습니다.

```sh
cp .env.example .env
```

배포 빌드는 `.env.production.local` 또는 CI secret으로 값을 넣습니다.

```sh
MAIN_VITE_MICROSOFT_CLIENT_ID=00000000-0000-0000-0000-000000000000
MAIN_VITE_MICROSOFT_TENANT_ID=consumers
```

`.env`, `.env.local`, `.env.production.local` 같은 실제 값 파일은 `.gitignore`로 제외합니다. 저장소에는 `.env.example`만 유지합니다.

앱 등록에서는 웹 로그인용 데스크톱 설정이 필요합니다.

- 플랫폼: Mobile and desktop applications
- Redirect URI: `http://localhost`
- Allow public client flows: Yes
- 권한 범위: `User.Read`, `Files.ReadWrite`, `offline_access`

개인 Microsoft 계정만 지원하는 앱 등록이라면 Tenant는 `consumers`를 사용합니다. Tenant를 `common`으로 쓰려면 앱 등록의 Supported account types가 개인 Microsoft 계정과 조직 계정을 모두 허용하도록 설정되어 있어야 합니다.

`redirect_uri is not valid` 오류가 나오면 `http://localhost`가 `Web` 플랫폼이 아니라 `Mobile and desktop applications` 플랫폼에 등록되어 있는지 확인합니다. MSAL Node는 로그인 중 임시 포트를 붙인 `http://localhost:<port>`를 사용하지만, Entra의 localhost redirect URI 매칭에서는 포트가 무시됩니다.

## 구조

```text
src/main/       Electron main process, 인증, Graph 호출, 플랫폼 경계 로직
src/preload/    renderer에 노출되는 타입 지정 IPC bridge
src/renderer/   원격 파일 관리자 UI
src/shared/     main, preload, renderer가 공유하는 타입
```

파일 목록은 Microsoft Graph `drive/root/delta`로 로컬 메타데이터 인덱스를 구성하며 탐색합니다. 앱 시작 또는 로그인 직후 탐색 인덱스를 백그라운드에서 구성하고, 폴더 이동은 로컬 인덱스를 우선 사용합니다. 아직 인덱스에 없는 폴더는 해당 폴더만 즉시 조회해 탐색을 막지 않습니다. 새로고침은 저장된 delta link로 변경분만 반영합니다.

파일 작업은 사용자가 명시적으로 선택한 항목에만 실행합니다. 업로드는 로컬 파일 선택 다이얼로그로 고른 파일을 현재 OneDrive 폴더에 올리고, 다운로드는 저장 위치 선택 다이얼로그로 지정한 경로에 저장합니다. 삭제는 OneDrive 휴지통으로 이동합니다.

업로드와 다운로드는 전송 작업으로 저장하며 진행률과 초당 전송 속도를 앱 안에서 표시합니다. 실행 중인 전송은 중지할 수 있고, 업로드는 Microsoft Graph upload session의 `nextExpectedRanges`를 기준으로 이어올리며 다운로드는 임시 파일과 HTTP Range 요청으로 이어받습니다. 전송 기록을 삭제하면 해당 upload session과 임시 다운로드 파일도 함께 제거해서 이어받기/이어올리기 상태를 폐기합니다.
