### Documentation is included in the Documentation folder ###
외교부 월별 세출집행현황 데이터를 추출하여 결과 파일 작성
외교부 홈페이지(https://www.mofa.go.kr/) - 정보공개 – 세입세출예산운용상황 에서 DataTable을 추출

### REFrameWork Template ###
**Robotic Enterprise Framework**
* Keeps external settings in *Config.xlsx* file and Orchestrator assets
사용자가 원하는 조회시작연월, 조회종료연월의 값을 받아서 데이터추출을 하기 위해
Data\Config.xlsx파일에 StartYYMM(조회시작연월), EndYYMM(조회종료연월) setting
*주의: Invoke VBA가 작동하기 위해서는 사전에 엑셀 환경설정필요(매크로활성화, 보안설정 등)

작업시작 전,  QueItem을 사용하지 않을 것이므로
1. TransactionItem을 String으로 변경 (TransactionItem을 엑셀시트 List(String)로 사용)
2. QueItem의 상태값 변경 부분 커맨드아웃처리

### How It Works ###

1. **INITIALIZE PROCESS**
1) Invoke KillAllProcesses workflow> Kill Process (크롬, 엑셀)
2) Invoke InitiAllApplications workflow> Open Browser (크롬)
3) Config의 조회연월값(StartYYMM, EndYYMM)을 활용하여 템플릿 시트를 생성(build data table)
   템플릿 시트를 Get Workbook Sheets의 output변수 list<String>으로 받아서 transactionItem으로 사용
   (TransactionItem은 datarow가 아닌 string으로 함)

2. **GET TRANSACTION DATA**
Invoke GetTransactionData workflow> Try 영역 
IF 조건값 TransactionNumeber가 list.Count(템플릿 시트 수)보다 작거나 같을 때까지 반복

3. **PROCESS TRANSACTION**
1) Invoke Process workflow> Config의 조회연월 값(StartYYMM, EndYYMM)을 검색하여 데이터추출 
2) 당월집행내역이 0원인 경우는 필터로 걸러내고 임시파일로 저장 (test.xlsx)

4. **END PROCESS**
1) 임시파일 Copy하여 결과파일로 복사, VBA 스타일 적용
2) 결과파일 내 종합시트 추가 (각 시트의 집행액 합계 등 요약) VBA 스타일 적용
3) 임시파일 삭제
4) Invoke CloseAllApplications workflow> CloseTab(크롬)

### For New Project ###

1. Check the Config.xlsx file and add/customize any required fields and values
2. Implement InitiAllApplications.xaml and CloseAllApplicatoins.xaml workflows, linking them in the Config.xlsx fields
3. Implement GetTransactionData.xaml and SetTransactionStatus.xaml according to the transaction type being used (Orchestrator queues by default)
4. Implement Process.xaml workflow and invoke other workflows related to the process being automated
