# SecureCRT-Report
SecureCRT 기반 Cisco Switch 보고서 자동 작성 프로그램입니다.  
문자열을 찾아 정보를 가져오는 방식이므로 장치 버전에 따라 정보를 가져오지 못하거나 부정확한 정보를 가져올 수 있습니다.
> 보고서 포함 내용
> * Model
> * Hostname
> * Version
> * Uptime
> * CPU
> * Memory(total/used)
> * Flash(total/free)
> * Temperature
> * Power Status
> * Fantray Status  

실행 환경
---------
* **SecureCRT 9.0 이상**  
  * 이전 버전에서는 python3 실행이 불가능할 수 있습니다.
* **Python3**
  * 3.9.13 버전 설치를 권장합니다.  
  (최신 버전의 python3가 필요하나 SecureCRT에서 최신 버전이 동작하지 않는 경우가 있습니다.)  
  * 파이썬 설치 시 *Add python.exe to PATH* 를 꼭 체크해주세요.
* **pandas, openpyxl**
  * 엑셀 데이터 추출 및 엑셀 저장을 위한 라이브러리입니다.
  * Windows Powershell에서 설치가능합니다.  
    ``` sh
    pip install pandas openpyxl
    ```
    
사용 방법
---------
* 접속 정보가 담긴 엑셀 준비  
  (host, port, username, password 헤더 포함 필수)  
1. SecureCRT 에서 Script > Run 클릭, writeReport.py 선택
2. Open File 창이 열리면 접속 정보 엑셀 파일 선택
3. 보고서 작성 및 txt 파일 생성  
   (접속 정보 엑셀 파일 경로에 hostname_config.txt, hostname_log.txt로 저장)
4. Save As 창이 열리면 저장 파일명을 입력  
   (취소 클릭 시 접속 정보 엑셀 파일 경로에 "report_YYYYMMDDHH24MISS.xlsx" 형식으로 저장)
5. 보고서 작성 완료
