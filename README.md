# auto_trade
개인프로젝트_전략에 따른 자동거래 시스템 개발


______________________
# 버전체크
아나콘다 32bit `4.10.1`

# 필요 라이브러리
pip install pykiwoom
____________________




실시간 거래 사용 파일
1. pytrader_2.py (ui부분 코드)
2. Kiwoom_2.py (키움증권 연결부분 코드)
3. pytrader.ui (ui 디자인)
--------------------------------------

## 실행순서
1. pytrader_2.py 파일 실행


![image](https://user-images.githubusercontent.com/70958560/158515592-498d79fb-d56f-44ec-9678-ce4c8fb86208.png)

-> 로그인 

2. 종목칸에 원하는 거래 종목 번호 입력 (ex. 코스피200선물 : 101S6000)
*참고 : 선물 거래만 가능


![image](https://user-images.githubusercontent.com/70958560/158515816-a1fc1be8-7c19-4522-b24f-65755d556aab.png)



3. 강제 청산시간, 강제청산기준(ex. 0.05) , 틱 단위(거래하는 기준 0.05) , 기준시간(프로그램시작할 때의 기준점 (9로 설정하면 싯가 기준 적용)   입력 후 프로그램 시작 클릭


![image](https://user-images.githubusercontent.com/70958560/158516053-7d2c8f55-8d10-4b10-9fb6-cc71cbed601a.png)



4. 오른쪽의 시가 리스트 생성 버튼 클릭


![image](https://user-images.githubusercontent.com/70958560/158516090-dd93dcbf-644a-40a9-8e3b-3fbf33408a55.png)





# 전체 화면

![image](https://user-images.githubusercontent.com/70958560/158516181-5badd0e5-2707-4d48-b8c7-5389132311fa.png)


