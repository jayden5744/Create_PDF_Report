# Create_PDF_Report 사용법


## 1. CLI 사용법
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
 
 ![CLI예제](/PDF/img/CLI예제.png)
 - do you want to continue? : 같은 실험의 다른 파일을 변환하고 싶으면 y, 마치고 싶으면 n
 - Enter the following file name : 다음 엑셀파일의 이름
 - Enter the following title : 다음 파일 실험보고서에 제목으로 들어가게 될 내용
 - Enter the following description : 다음 파일 실험보고서에 시험설명으로 들어가게 될 내용
 
### CLI에서 GUI 실행
- python excel2pdf.py --gui


## 2. GUI 사용법
