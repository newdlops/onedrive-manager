# 1차 목표

## 목표

사용자가 Windows와 macOS 앱에서 Microsoft 계정으로 OneDrive에 접속하고, 계정의 OneDrive 파일과 폴더를 원격 탐색기 형식으로 확인할 수 있게 만든다.

이 앱의 제품 정체성은 동기화 클라이언트가 아니라 OneDrive 원격 파일 관리자다.

## 완료 기준

- Microsoft 계정 로그인과 로그아웃을 지원한다.
- 현재 로그인한 계정의 이름과 이메일을 앱 안에서 확인할 수 있다.
- OneDrive 루트 폴더의 파일과 폴더 목록을 불러온다.
- 폴더를 더블클릭하거나 선택해서 하위 폴더로 이동할 수 있다.
- 탐색기 목록은 이름, 유형, 크기, 수정일을 표시한다.
- 로딩, 빈 폴더, 인증 만료, 권한 부족, 네트워크 실패 상태를 한국어로 표시한다.
- Windows와 macOS에서 같은 코드 경로를 우선 사용한다.

## 현재 범위에서 제외

- 파일 이동
- 공유 링크 생성과 권한 관리
- 다중 계정 동시 로그인
- 자동 동기화, 실시간 변경 감지, 백그라운드 동기화
- 로컬 OneDrive 동기화 상태 제어

## 이후 보강

- 전송 취소와 재시도 정책을 정교하게 만든다.
- 파일 이동과 복사 같은 원격 파일 관리 명령을 추가한다.
- 로컬 폴더와 OneDrive 폴더를 계속 맞추는 동기화 기능은 제공하지 않는다.

## 구현 방향

- 인증과 Microsoft Graph 호출은 Electron main process에서 처리한다.
- renderer는 preload를 통해 타입이 지정된 IPC만 호출한다.
- OneDrive 파일 목록은 Microsoft Graph Drive API를 기준으로 가져온다.
- OAuth 토큰 캐시는 Electron `safeStorage`로 암호화해서 앱 사용자 데이터 경로에 저장한다.
- 원격 파일 관리 기능은 Graph API 작업으로 구현하고, 로컬 파일시스템 작업은 사용자가 선택한 다운로드/업로드 대상에만 제한한다.
- Windows 사용자가 많다는 전제를 유지해서 경로, 파일명, 날짜, 네트워크 오류 처리를 Windows 기준으로도 확인한다.

## 현재 구현

- MSAL Node interactive flow로 기본 브라우저에서 Microsoft 웹 로그인을 진행한다.
- 로그인 응답은 loopback redirect URI로 받아서 앱에서 토큰으로 교환한다.
- Microsoft Entra Application client ID는 빌드 환경에서 주입하고, 누락된 개발 환경에서만 앱 안의 로그인 설정을 표시한다.
- 로그인 후 Microsoft Graph `drive/root/delta`로 OneDrive 메타데이터 인덱스를 백그라운드에서 구성한다.
- 인덱싱 중에도 탐색은 계속 가능해야 하며, 아직 인덱스에 없는 폴더는 해당 폴더만 즉시 조회해서 표시한다.
- 탐색기 목록은 로컬 인덱스를 기준으로 이름, 유형, 크기, 수정일을 표시한다.
- 새로고침은 저장된 `@odata.deltaLink`로 변경분만 반영한다.
- 선택한 항목의 이름 변경과 삭제를 지원한다. 삭제는 OneDrive 휴지통 이동으로 처리한다.
- 현재 폴더에 로컬 파일을 업로드하고, 선택한 파일을 사용자가 지정한 로컬 경로로 다운로드한다.
- 업로드와 다운로드 진행률 및 초당 전송 속도를 표시하고, 실행 중인 전송은 중지할 수 있다.
- 중지되거나 중단된 전송은 저장된 전송 작업에서 재개할 수 있다.
- 전송 기록 삭제 시 저장된 이어받기/이어올리기 정보와 임시 파일도 함께 삭제한다.

## 필요한 설정

- Microsoft Entra 앱 등록
- 데스크톱 public client용 client ID
- Mobile and desktop applications 플랫폼의 `http://localhost` redirect URI
- public client flow 허용
- 기본 테넌트 값은 개인 Microsoft 계정용 `consumers`
- 권한 범위는 `User.Read`, `Files.ReadWrite`, `offline_access`

앱 등록이 Personal Microsoft accounts only라면 Tenant는 `consumers`를 사용한다. Tenant를 `common`으로 쓰려면 앱 등록의 Supported account types가 개인 Microsoft 계정과 조직 계정을 모두 허용해야 한다.

`http://localhost`는 `Web` 플랫폼이 아니라 `Mobile and desktop applications` 플랫폼에 등록한다. MSAL Node의 interactive login은 loopback 서버를 열면서 임시 포트를 사용하지만, Microsoft Entra의 localhost redirect URI 매칭에서는 포트가 무시된다.
