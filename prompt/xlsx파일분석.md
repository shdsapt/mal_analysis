- 다음은 .xlsx, xls 등 엑셀 파일을 분석하는 단계에 대한 설명이고 아래에 명시된 명령을 단계별로 실행해줘

1. 1. 내가 요청한 파일에 대해 file 명령어 실행하고 그 결과를 출력 및 해석 
   - ex) file suspicious.png
2. olevba 명령어를 실행하여 결과를 출력하고 매크로가 있는지 여부를 분석해줘
   - ex) olevba suspicious.xlsx
3. 마지막으로 strings 명령어를 실행하여 결과를 출력하고 악의적인 링크 및 파일 포함 여부를 확인해줘
   - ex) strings suspicious.xlsx | grep -iE 'http|ftp|tcp|udp|powershell|cmd.exe|vbs|exe|bat'

4. 이외 결과에 대해 추가적으로 분석할 방안이 있다면 제안해줘 
