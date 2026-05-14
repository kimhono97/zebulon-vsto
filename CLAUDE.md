# Claude Code Instructions

## 현재 상태: 모노레포 분리 직후

이 리파지토리는 `kimhono97/jym_ppt`에서 분리된 직후 상태입니다.
**일반 개발 가이드 이전에 \"분리 후속 작업\"을 먼저 진행해야 합니다.**

이 프로젝트는 독립 실행되는 Visual Studio Tools for Office (VSTO) PowerPoint Add-in으로, 모노레포 시절에도 외부 의존이 거의 없었습니다. 따라서 후속 작업은 \"빌드 환경 검증과 문서 보완\" 위주이며 다른 두 폴리레포(`zebulon-provider`, `zebulon-exporter`)에 비해 가벼운 편입니다.

## 세션 시작 시 작업 순서

1. **`MIGRATION_NOTES.md`를 먼저 읽기** — 분리 배경, 보존/손실된 것, 후속 작업 체크리스트 확인
2. 체크리스트의 각 항목을 순서대로 진행. 항목마다 완료되면 `- [x]`로 체크 표시
3. 외부 도구 검증(Visual Studio 빌드, PowerPoint add-in 동작 확인)은 사용자에게 안내하고 결과를 받아 다음 단계 진행
4. **이 CLAUDE.md 자체는 모든 후속 작업이 완료되기 전까지 수정하지 말 것** (마이그레이션 흐름 보존)
5. 코드 변경은 체크리스트 항목 수행에 필요한 최소 범위로 한정
6. 의심스러우면 사용자에게 확인

## 정리 단계 (모든 후속 작업 체크리스트가 완료된 후에만 수행)

`MIGRATION_NOTES.md`의 모든 체크박스가 `[x]`가 되었는지 확인한 뒤:

1. **`MIGRATION_NOTES.md` 삭제**
2. **`readme.md` 상단의 migration 배너 한 줄 제거**
3. **이 `CLAUDE.md`를 일반 프로젝트 가이드로 재구성**. 다음 섹션을 포함:
   - 프로젝트 개요 (PowerPoint VSTO Add-in으로, UDP 기반 sync 기능 등)
   - 빌드 / 디버깅 (Visual Studio 2022, NuGet 복원, F5 디버그 등)
   - 주요 디렉토리·파일 설명 (`ZebulonVSTO/` 솔루션 구조)
   - 배포 방식 (VSTO 배포 매니페스트, ClickOnce 등 — 해당하는 경우)
   - 의존성 (NuGet 패키지, Office/.NET 버전 등)
4. 최종 점검: 정리 commit으로 마이그레이션 완료를 명확히 마킹
   - 예시 commit message: `chore: complete polyrepo migration, regenerate CLAUDE.md`

## 주의사항

- VSTO 빌드 산출물(`bin/`, `obj/`, `*.user`)을 실수로 commit하지 않도록 주의
- 빌드 환경(VS workload, .NET Framework 버전)이 정확히 맞아야 add-in 동작이 보장됨
