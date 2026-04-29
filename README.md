# 문서 일괄 암호 설정기

개인정보가 포함된 문서 파일에 공통 열기 암호를 일괄 적용하기 위한 Windows용 도구입니다.

HWP/HWPX, PDF, Excel, Word, PowerPoint 문서를 검색하고, 지원되는 형식에 암호를 설정합니다. 이미지 파일은 암호화 대상이 아니며 삭제 후보 점검 또는 목록 정리용으로만 다룹니다.

## 주요 기능

- 하위 폴더 포함 파일 검색
- HWP/HWPX 한글 COM 자동화 암호 설정
- PDF 열기 암호 설정
- XLSX/XLSM 열기 암호 설정
- DOCX/DOCM, PPTX/PPTM 열기 암호 설정
- 이미 암호가 있는 파일 자동 스킵
- 그림파일 목록 제거
- 파일 유형별 선택
- 다음 100개씩 선택
- 실패 파일만 재선택
- 실행 전 점검
- 삭제 후보 점검
- 작업 로그 저장

## 지원 형식

| 구분 | 확장자 | 비고 |
| --- | --- | --- |
| 한글 | `.hwp`, `.hwpx` | 한컴오피스 한글 COM 필요 |
| PDF | `.pdf` | `pikepdf` 필요 |
| Excel | `.xlsx`, `.xlsm` | `msoffcrypto-tool` 필요 |
| Word | `.docx`, `.docm` | Office Open XML 기반 |
| PowerPoint | `.pptx`, `.pptm` | Office Open XML 기반 |
| 구형 Office | `.xls`, `.doc`, `.ppt` | 자동 열기 암호 설정 제한 |
| 이미지 | `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.tif`, `.tiff`, `.webp` | 암호화 대상 아님 |

## 설치

Python 3.10 이상 권장.

```powershell
python -m pip install pywin32 pikepdf openpyxl xlrd msoffcrypto-tool tkinterdnd2 pyinstaller
```

`tkinterdnd2`는 드래그앤드롭 기능용입니다. 설치되어 있지 않아도 프로그램 실행은 가능하지만 드래그앤드롭은 비활성화됩니다.

설치 확인:

```powershell
python -c "import tkinterdnd2; print('tkinterdnd2 OK')"
python -c "import pikepdf, openpyxl, msoffcrypto; print('document libs OK')"
```

## 실행

소스 파일로 실행:

```powershell
python "문서일괄암호설정기_v1.41.py"
```

## EXE 빌드

아이콘 파일 `문서일괄암호설정기.ico`를 사용할 때:

```powershell
pyinstaller --noconfirm --clean --onefile --windowed --icon "문서일괄암호설정기.ico" --name "문서일괄암호설정기_v1.42" --hidden-import tkinterdnd2 --collect-data tkinterdnd2 "문서일괄암호설정기_v1.41.py"
```

빌드 결과:

```text
dist\문서일괄암호설정기_v1.42.exe
```

기존 `.spec` 파일이 오래된 소스 파일을 가리킬 수 있으므로, 배포용 빌드는 위처럼 최신 `.py` 파일을 직접 지정하는 방식을 권장합니다.

## 권장 사용 순서

1. 원본이 아닌 사본 폴더에서 먼저 테스트합니다.
2. 대상 폴더를 선택하고 파일을 검색합니다.
3. 필요하면 `그림파일 목록 제거`를 먼저 실행합니다.
4. `파일 유형별 선택`으로 먼저 처리할 형식을 고릅니다.
5. HWP/HWPX는 `다음 100개 선택`으로 나눠 처리하는 것을 권장합니다.
6. `암호 파일 스킵` 체크를 유지한 상태로 암호 설정을 실행합니다.
7. 완료 후 성공 파일 중 일부를 직접 열어 암호 적용 여부를 확인합니다.
8. 실패 파일은 `실패 파일만 재선택`으로 다시 처리합니다.

## 대량 파일 처리 안내

파일이 수천 개 이상인 경우 `암호 파일 목록 제거`는 오래 걸릴 수 있습니다. 이 기능은 모든 파일의 암호 여부를 미리 검사하기 때문입니다.

대량 작업에서는 `암호 파일 목록 제거`를 필수로 누르지 않아도 됩니다. `암호 파일 스킵`이 체크되어 있으면 암호 설정 실행 중 이미 암호가 있는 파일은 자동으로 건너뜁니다.

권장 흐름:

- 이미지가 많으면 `그림파일 목록 제거`
- 문서 유형별로 선택
- HWP/HWPX는 100개 이하로 나눠 실행
- 실패 파일만 따로 재시도

## 주의사항

- 작업 전 백업을 권장합니다.
- 한글 문서는 한컴오피스 설치 및 COM 자동화 환경에 영향을 받습니다.
- HWP/HWPX 대량 처리 중에는 마우스/키보드 조작을 피하는 것이 좋습니다.
- `.xls`, `.doc`, `.ppt` 같은 구형 Office 형식은 자동 열기 암호 설정이 제한됩니다.
- Excel 암호 설정에는 `msoffcrypto-tool`이 필요합니다. 시트 보호는 파일 열기 암호가 아니므로 성공으로 처리하지 않습니다.
- 개인정보 파일 처리 후에는 암호가 없는 원본 또는 백업 파일이 남아 있지 않은지 확인해야 합니다.

## 문제 해결

### 드래그앤드롭이 안 되는 경우

빌드 시 아래 옵션이 포함되어야 합니다.

```powershell
--hidden-import tkinterdnd2 --collect-data tkinterdnd2
```

또한 빌드에 사용한 Python 환경에 `tkinterdnd2`가 설치되어 있어야 합니다.

```powershell
python -c "import tkinterdnd2; print(tkinterdnd2.__file__)"
```

### HWP/HWPX 처리 중 실패가 많은 경우

- 한컴오피스 한글을 완전히 종료한 뒤 다시 실행합니다.
- HWP/HWPX를 100개 이하로 나눠 처리합니다.
- 큰 문서나 그림이 많은 문서는 별도 배치로 처리합니다.
- 실패 파일만 재선택해 재시도합니다.

### Excel 암호 설정 실패

`msoffcrypto-tool` 설치 여부를 확인합니다.

```powershell
python -m pip install msoffcrypto-tool
```

`.xls` 파일은 자동 열기 암호 설정이 제한되므로 `.xlsx`로 변환 후 처리하는 것을 권장합니다.

## 배포 전 체크리스트

- 정보창 버전과 파일명 버전이 일치하는지 확인
- 최신 `.py` 파일 기준으로 exe 빌드
- 드래그앤드롭 동작 확인
- PDF, Excel, HWP/HWPX 샘플 암호 적용 확인
- 실패 로그 저장 확인
- 백업 폴더 또는 원본 파일 개인정보 잔존 여부 확인
