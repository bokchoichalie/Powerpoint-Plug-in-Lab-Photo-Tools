# Lab Photo Tools 0.1.17 Preview

## 설치 파일 하나로 설치

**[LabPhotoTools-0.1.17-Setup.exe 다운로드](https://github.com/bokchoichalie/Powerpoint-Plug-in-Lab-Photo-Tools/releases/download/v0.1.17/LabPhotoTools-0.1.17-Setup.exe)**

1. PowerPoint를 닫습니다.
2. 위 EXE를 실행하고 **설치**를 누릅니다.
3. 설치가 끝나면 PowerPoint에서 **Lab Photo Tools** 탭을 엽니다.

Python을 따로 설치하거나 ZIP을 풀 필요가 없습니다. 처음 한 번 인터넷으로 전용 Python, 사진 처리 라이브러리와 배경 제거 모델을 자동 준비합니다. 초기 다운로드 약 220 MB, 설치 공간 약 1 GB입니다. 이후 사진 처리는 오프라인에서 동작합니다. 친구에게 위 설치 파일이나 이 릴리스 링크를 공유하면 됩니다.

## 변경 사항

- Python 공식 임베디드 배포판을 프로그램 전용 폴더에 설치합니다. 기존 Python이나 PATH 설정에 의존하지 않습니다.
- Microsoft Visual C++ 런타임이 없으면 공식 설치 파일의 서명을 확인하고 설치합니다. 이 단계에서 Windows 관리자 승인이 필요할 수 있습니다.
- 설치 진행 안내·로그·실패 후 재시도 기능과 Windows 설정의 앱 제거 항목을 제공합니다.
- 기존 사용자의 정상 엔진과 번호 서식 설정을 보존합니다. 0.1.16의 라벨 상자 서식, 알파벳) 라벨, 그룹 사진 측정 복사, n점원·점편집 기능을 포함합니다.

## 환경 및 검증

Intel/AMD 64비트 Windows 10/11, .NET Framework 4.8 이상, Windows 데스크톱 PowerPoint가 필요합니다. PowerPoint는 별도로 설치되어 있어야 합니다. 32비트·64비트 PowerPoint에 등록하며 macOS·웹 PowerPoint·ARM64 Windows는 지원 대상이 아닙니다.

기존 Python을 참조하지 않는 별도 런타임 설치와 사진 처리, EXE 업데이트, 기존 설정 보존, 실제 PowerPoint 추가 기능 로드를 검사했습니다. 초기화된 새 노트북 전체 환경이나 Visual C++ 관리자 설치 단계는 실기 검증하지 않았습니다. 자세한 기록은 저장소의 `VALIDATION.md`에 있습니다.

ZIP은 수동 설치·기존 방식 업데이트용으로 함께 제공합니다. 일반 설치에는 **Setup.exe**를 사용하세요.
