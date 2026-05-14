# Polyrepo Migration Notes (zebulon-vsto)

> 이 문서는 이 리파지토리가 모노레포(`kimhono97/jym_ppt`)에서 분리된 직후의 맥락과 분리 후 남은 작업을 기록한 것입니다.
> **모든 후속 작업이 완료되면 이 파일은 삭제됩니다.** 분리 사실은 git history와 README 한 줄로만 남게 됩니다.

## 출처 및 분리 시점

- 원본 모노레포: [`kimhono97/jym_ppt`](https://github.com/kimhono97/jym_ppt)
- 원본 경로: `utils/ZebulonVSTO/`
- 분리 일자: 2026-05-14
- 분리 방식: `git filter-repo --path utils/ZebulonVSTO/ --path-rename utils/ZebulonVSTO/:`

## 분리 배경

- `jym_ppt`는 가사 데이터 + 다수 코드 프로젝트가 섞인 모노레포로 비대해짐
- 팀원은 `jym_ppt`에 collaborator로 등록되어 가사 편집만 하는데, 코드 폴더까지 노출되는 부담
- Vercel Hobby 플랜의 \"non-owner collaborator push 시 deploy 안 됨\" 제약 → 코드 리파지토리는 owner 단독 push 구조로 분리할 필요
- 최종 그림: `jym_ppt`는 가사 전용 데이터 리파지토리(추후 private 복귀), 코드는 각각 polyrepo에서 owner가 독립 운영

## 보존된 것 / 손실된 것

- ✅ 보존: `utils/ZebulonVSTO/`를 건드린 6개 commit, 파일 변경 이력 전부
- ⚠️ 손실(의도적): 머지 commit 1개 (rebase 기본 동작; **코드 내용 손실 없음 검증 완료**)
- ⚠️ 손실(의도적): 가사·다른 코드 폴더의 commit 전부, 원본 SHA(재작성됨)
- 📌 추가: GitHub 자동 생성 `Initial commit (LICENSE)`이 history 최하단에 있음 (rebase로 통합)

## 분리 직후 상태

- 기본 브랜치: `main` (LICENSE init + 6개 VSTO commit)
- 모노레포 시절 `CLAUDE.md`는 존재하지 않았음 (이 리파지토리는 처음부터 새 CLAUDE.md를 만드는 케이스)

## 후속 작업 체크리스트

### A. 빌드·실행 환경 검증

- [ ] Visual Studio 2022 (Community v17.5.2+)에서 `ZebulonVSTO.sln`을 그대로 열어 빌드되는지 확인
- [ ] 필요한 워크로드/패키지가 모두 설치되어 있는지 검증 (Office/SharePoint 개발, VSTO, .NET Framework 4.7.2/4.8)
- [ ] NuGet packages가 정상 복원되는지 확인 (`Microsoft.Bcl.AsyncInterfaces`, `System.Text.Json` 등 — readme 참고)
- [ ] PowerPoint 2013+에서 add-in 정상 동작 확인 (Sync 기능: `alert`, `select`, `showslide`, `hideslide`)

### B. 문서·메타 정리

- [ ] `readme.md`에 빌드/실행 가이드가 충분한지 검토. 필요 시 보완 (배포 방법, 디버깅 팁 등)
- [ ] (선택) 첫 release/tag 생성으로 분리 시점을 마킹 (예: `v1.0-polyrepo`)
- [ ] (선택) `.gitignore` 점검 — `bin/`, `obj/`, `*.user` 등 VS 산출물 제외 여부

### C. 정리 (체크리스트 전부 끝난 뒤)

- [ ] 이 `MIGRATION_NOTES.md` 파일 삭제
- [ ] `readme.md` 상단의 migration 배너 한 줄 제거
- [ ] `CLAUDE.md`를 일반 프로젝트 가이드로 재구성 (본 CLAUDE.md의 \"정리 단계\" 안내 참고)
