ZebulonVSTO - PowerPoint 슬라이드 동기화 추가기능
==================================================

UDP로 한 PowerPoint(송신)의 슬라이드 이동을 다른 PowerPoint(수신)에 동기화하는
추가기능입니다.

[ 요구사항 ]
- Windows
- PowerPoint 2013 이상
- VSTO 2010 런타임 (Microsoft Visual Studio 2010 Tools for Office Runtime)
  없으면 설치 시 안내됩니다. (microsoft.com/download, id 48217)

[ 설치 ]  ※ 관리자 권한 불필요 (현재 사용자 전용)
1. 이 압축을 폴더에 모두 풉니다.
2. PowerPoint를 닫습니다.
3. Install.ps1 을 PowerShell로 실행합니다.
   - 탐색기에서 마우스 오른쪽 > "PowerShell에서 실행", 또는
   - PowerShell 창에서:  powershell -ExecutionPolicy Bypass -File .\Install.ps1
4. 설치 중 인증서 신뢰 확인창이 뜨면 "예"를 클릭합니다. (자체서명 인증서 1회 신뢰)
5. 끝나면 결과 창(설치 완료/실패)이 뜹니다. "INSTALL COMPLETE"면 성공입니다.
6. PowerPoint를 다시 시작하면 추가 기능 탭에 "Zebulon"이 나타납니다.

   ※ 제거(Uninstall.ps1)도 동일하게 끝에 결과 창이 뜹니다.

[ 업데이트 ]
새 압축을 받아 같은 방식으로 Install.ps1 을 다시 실행하면 덮어쓰기 됩니다.
(PowerPoint는 닫은 상태에서)

[ 제거 ]
PowerPoint를 닫고 Uninstall.ps1 을 실행합니다.

[ 사용법 ]
- PowerPoint > 추가 기능 탭 > "Zebulon" 그룹 > [동기화 시작] 클릭
- 설정 마법사가 열립니다:
  1) 역할 선택: 송신(발표자) 또는 수신(청중)
  2) 수신: 수신 포트만 정하고 [시작] (기본 8291)
  3) 송신: 설정 방식 선택
     - 자동으로 찾기: 같은 네트워크의 수신기를 검색 → 목록에서 선택
       (또는 "모든 수신기로 전송"=브로드캐스트), 포트 지정 후 [시작]
     - 직접 입력: 원격 IP/포트를 직접 입력 후 [시작]
- 작동 중에는 리본에 자신/상대의 IP·포트와 상태가 표시됩니다.
- [동기화 중지]로 종료, [콘솔]로 주고받는 트래픽 로그 확인
- 명령: alert <글자>, select <번호>, showslide <번호>, hideslide

  ※ "자동으로 찾기"는 송신·수신 양쪽이 모두 이 버전(또는 이후)이고 같은 서브넷일
     때만 동작합니다. 브로드캐스트가 차단된 망에서는 "직접 입력"을 사용하세요.
     (탐색용 UDP 포트 8290을 함께 사용합니다.)

[ 진단 도구 (Tools 폴더) ]
설치된 추가기능이 잘 동작하는지 확인하는 PowerShell 도구입니다. (선택)
- 설치된 추가기능을 "수신" 모드로 시작한 뒤, 명령이 들어가는지 확인:
    Tools\Send-SyncCommand.ps1 -Command "select 2" -Port 8291
- 추가기능을 "송신" 모드(원격 포트 8292)로 두고, 무엇을 보내는지 관찰:
    Tools\Start-SyncSession.ps1 -Mode Receiver -Port 8292
  ※ 같은 PC에서는 추가기능이 8291을 점유하므로, 진단 도구는 다른 포트를 쓰세요.
- "자동으로 찾기" 목록에 잡히는지 확인 (수신기 흉내내기):
    Tools\Start-SyncSession.ps1 -Mode Responder -Port 8291

[ 문제 해결 ]
- "Zebulon" 탭이 안 보임:
    PowerPoint 완전 종료 후 재실행. 그래도 없으면 VSTO 런타임 설치 여부 확인.
    파일 > 옵션 > 추가 기능 > 관리: COM 추가 기능 에서 사용 안 함으로 빠졌는지 확인.
- "알 수 없는 게시자" 경고:
    Install.ps1 의 인증서 신뢰 단계가 끝나지 않았을 수 있음. 다시 실행.
- 동기화가 안 됨:
    방화벽이 UDP 포트(기본 8291)를 막는지 확인. 송신/수신 포트가 맞는지 확인.

문의/소스: https://github.com/kimhono97/zebulon-vsto
