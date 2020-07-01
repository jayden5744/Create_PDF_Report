# Create_PDF_Report 사용법

## 1. Requirements
- Python 2.7
- docx 0.2.4
- python-docx 0.8.10
- numpy 1.16.6
- pandas 0.24.2
- openpyxl 2.3.0
- matplotlib 2.2.4
- xlrd 1.2.0
- comtypes 1.1.7

## Install
``` Python
 pip install -r requirements.txt
```
---
## 2. CLI 사용법
### 필수 command
- 파일종류 : 해당 실험의 종류 
  - ex) --sa12, --sa15, --sa10_1
- filename : 엑셀파일의 이름 
  - ex) --filename "SA12"

### 옵션 command
- title : 실험보고서에 제목으로 들어가게 될 내용 
  - ex) --title "SA12 Specified Powr Factor 기능 시험"
- description : 실험보고서에 시험설명으로 들어가게 될 내용 
  - ex) --description : "다음 시험은 이러이러하다"
- path : 해당 엑셀파일이 있는 폴더의 경로
  - ex) --path "C://Users"
- save_path : PDF의 저장경로
  - ex) --save_path "C://Users/save_folder"
  
### 예제
 - python excel2pdf.py --sa12 --filename "SA12"
 - python excel2pdf.py --sa15 --filename "SA15" --title "실험제목" --description "시험설명" --path "C://Users" --save_path "C://Users/save_folder"
 
 ![CLI예제](/img/CLI예제.png)
 - do you want to continue? : 같은 실험의 다른 파일을 변환하고 싶으면 y, 마치고 싶으면 n
 - Enter the following file name : 다음 엑셀파일의 이름
 - Enter the following title : 다음 파일 실험보고서에 제목으로 들어가게 될 내용
 - Enter the following description : 다음 파일 실험보고서에 시험설명으로 들어가게 될 내용
 
### CLI에서 GUI 실행
- python excel2pdf.py --gui


## 3. GUI 사용법
## 메인화면
![GUI메인](/img/GUI메인.png)

1. 파일경로 : 불러올 엑셀파일을 선택합니다.
2. 시험종류 선택 : 보고서 유형을 sa08/sa09/sa9_1/sa10/sa10_1/sa11/sa12/sa13/sa14/sa15 중 선택합니다.
3. 시험 제목 : 실험보고서에 제목으로 들어가게 될 내용을 입력합니다.
4. 시험 설명 : 실험보고서에 시험설명으로 들어가게 될 내용을 입력합니다.
5. 저장경로 설정 : PDF 파일을 저장할 폴더를 선택합니다.
6. 변환 실행 : 1~5번의 내용을 입력 후 클릭 시 PDF로 변환합니다.

## 세부화면
1. 메인화면 : 엑셀파일이 있는 폴더에 들어가서 해당 파일을 선택합니다.
![GUI_1](/img/GUI_1.png)

2. 시험종류 선택 : 해당되는 보고서 유형을 선택합니다.
![GUI_2](/img/GUI_2.png)

3. 저장경로 설정 : PDF파일을 저장할 경로를 선택합니다.
![GUI_3](/img/GUI_3.png)
