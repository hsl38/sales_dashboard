# sales_dashboard
generates sales dashboard from data stored in an xlsx

# 기능

-   세일즈 데이터가 들어있는 xlsx 파일 복사본 파일을 만든다.

    -   데이터 xlsx 파일 이름: sales_data_m.xlsx로 고정되어 있다.

    -   m: xlsx 파일의 버전 번호이다. 파이썬 스크립트 파일의 이름
        sales_dashboard_m.py의 m과 동일해야 한다.

    -   복사본 파일의 이름은 yyyymmdd_hhmmss_sales_data_m.xlsx 이다.

        -   yyyymmdd_hhmmss: 복사본이 만들어진 연월일_시분초 이다.

-   복사본 xlsx 파일의 데이터를 처리하여 세일즈 대시보드 차트들을
    만든다.

-   차트들을 html 파일로 출력한다.

    -   출력된 파일의 이름은 yyyymmdd_hhmmss_sales_dashboard_m.html

        -   yyyymmdd_hhmmss: 복사본이 만들어진 연월일_시분초 이다.

-   차트 데이터를 복사본 xlsx 파일의 지정 시트의 지정된 셀에 출력한다.

    -   xlsx 파일에는 지정된 위치의 데이터로 차트들을 출력하도록 사전
        설정이 되어 있다.

-   대시보드 출력 완료 후 원본 sales_data_m.xlsx 파일은
    yymmdd_hhmmss_sales_dashboard_m.xlsx 파일로 덮어써지는 방식으로
    업데이트 된다.

# 디렉토리 및 파일 구조

-   아래와 같다.

> \[세일즈 대시보드 디렉토리\]\\py
>
> \[세일즈 대시보드 디렉토리\]\\sales_data
>
> \[세일즈 대시보드 디렉토리\]\\sales_dashboard

-   py: 이 스크립트 파일(sales_dashboard_m.py)이 있는 디렉토리다.

-   sales_data: 데이터 파일 (sales_data_m.xlsx)이 있는 디렉토리다.

-   sales_dashboard: 데이터 파일의 복사본과 html 형식의 대시보드 파일이
    저장된다.

# 대시보드 항목

## Sales KPI

-   스크립트를 수행하는 연도와 월을 기준

-   매출 전망, SSO (Single Sales Opportunity), SSO당 매출

![A picture containing diagram Description automatically
generated](media/image1.png)

## 월별 누적 매출 전망

![Chart, bar chart Description automatically
generated](media/image2.png)

## 솔루션 필드별 매출 전망

![Chart, bar chart Description automatically
generated](media/image3.png)

## 솔루션별 매출 전망

![Chart, bar chart Description automatically
generated](media/image4.png)

## 고객사별 매출 전망

![Chart, bar chart Description automatically
generated](media/image5.png)

## 고객군별 매출 전망

![Chart, bar chart Description automatically
generated](media/image6.png)

## 사업개발팀 Top 5 SSO

![Chart, funnel chart Description automatically
generated](media/image7.png)

## 솔루션별 Top 3 SSO

![Chart, bar chart Description automatically
generated](media/image8.png)

## Booking KPI

-   sales는 매출 발생 연도가 당해 연도인 SSO의 매출이다. 실제 수주는
    이전 연도일 수도 있다.

-   booking은 수주가 당해 연도인 SSO의 매출이다. 실제 매출은 다음 연도나
    그 이후에 발생할 수도 있다.![A picture containing graphical user
    interface Description automatically
    generated](media/image9.png)
