### ZebulonVSTO (PowerPoint Add-in)

### I. IDE Settings

1. Visual Studio 2022 이상

    - Community v17.5.2 이상 (VS 2026 Community v18.1 에서 빌드 검증 완료)

    - Office/SharePoint 개발
    
    ```
    (포함됨) Visual Studio용 Office 개발자 도구
    (포함됨) .Net Framework 4.7.2 개발 도구
    (포함됨) Developer Analytics Tools
    (선택 사항) VSTO(Visual Studio Tools for Office)   ← 빌드에 필수. 누락 시 빌드 실패
    (선택 사항) 웹 배포
    (선택 사항) .Net Framework 4.8 개발 도구
    ```

    > ⚠️ "VSTO(Visual Studio Tools for Office)" 컴포넌트가 없으면
    > `Microsoft.VisualStudio.Tools.Office.targets`를 찾지 못해 빌드가 실패합니다.
    > Visual Studio Installer → 수정 → 개별 구성 요소에서 설치하세요.

2. PowerPoint 추가기능 VSTO Project

    - PowerPorint 2013+

    - 외부 NuGet 패키지 의존성 없음 (프레임워크 전용)

      JSON 직렬화는 .NET Framework 내장 `System.Runtime.Serialization`
      (`DataContractJsonSerializer`)을 사용합니다. 별도의 NuGet 복원이 필요 없습니다.


### II. Features

1. Sync (UDP Sender/Receiver)

    - `alert [text]` : Open an alert dialog with the text message.

    - `select [n]` : Select the n-th slide in the editor window.

    - `showslide [n]` : Start the slide show and move to the n-th slide.

    - `hideslide` : Finish the slide show.


### III. Build & Debug

1. Visual Studio (권장)

    - `ZebulonVSTO.sln`을 연다 → NuGet 패키지가 자동 복원된다.
    - **Debug** 또는 **Release** 구성으로 빌드 (`Any CPU`).
    - **F5**를 누르면 PowerPoint가 추가 기능이 등록된 상태로 실행된다 (런타임 테스트는 이 경로로 한다).

2. 명령줄 빌드 (CI / 빠른 검증용)

    ```bash
    # 빌드 (NuGet 복원 불필요 — 외부 패키지 의존성 없음)
    msbuild ZebulonVSTO.sln -p:Configuration=Release -p:VisualStudioVersion=10.0
    ```

    > 💡 명령줄 MSBuild는 `VisualStudioVersion`을 자동으로 최신 버전(예: 18.0)으로
    > 설정하지만, VSTO 빌드 타겟(`OfficeTools`)은 `v10.0` 경로에 설치됩니다.
    > 따라서 CLI 빌드 시 `-p:VisualStudioVersion=10.0`을 명시해야 합니다.
    > Visual Studio IDE(F5/빌드)에서는 이 설정이 자동 처리되므로 신경 쓸 필요 없습니다.

3. 단위 테스트

    순수 로직(명령 파서, 메시지 직렬화)에 대한 xUnit 테스트는 `tests/ZebulonVSTO.Tests`에 있습니다.

    ```bash
    dotnet test tests/ZebulonVSTO.Tests
    ```

    - 테스트 프로젝트는 SDK 스타일(`net472`)이며 `ZebulonVSTO.sln`에 포함되지 않습니다.
      VSTO/MSBuild 빌드와 독립적으로 `dotnet` 툴체인으로 빌드·실행됩니다.
    - PowerPoint/COM이 필요한 부분(`SyncManager` 통신, 슬라이드 동작)은 단위 테스트 대상이
      아니며, 위의 F5 방식으로 수동 검증합니다.


### IV. Deployment

- 빌드 결과물(`bin\<Configuration>\`)에는 ClickOnce / VSTO 배포 매니페스트가 포함됩니다:
  `ZebulonVSTO.dll`, `ZebulonVSTO.dll.manifest`, `ZebulonVSTO.vsto`.
- 매니페스트 서명에는 `ZebulonVSTO_TemporaryKey.pfx`(저장소에 포함된 임시 키)를 사용합니다.
  실제 배포 시에는 정식 코드 서명 인증서로 교체하세요.
  - 인증서가 사용자 인증서 저장소에 없으면 빌드 중 경고(MSB3327)가 표시되지만,
    DLL 컴파일 자체는 정상적으로 완료됩니다.
- 설치 대상 PC에는 **VSTO 런타임**(Visual Studio 2010 Tools for Office Runtime)이 필요합니다.