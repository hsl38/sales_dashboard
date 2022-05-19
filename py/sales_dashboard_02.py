# %%
'''
sales_dashboard_m.py

세일즈 데이터가 있는 sales_data_m.xlsx 파일에서
yyyymmdd_hhmmss_sales_dashboard_m.html과
yyyymmdd_hhmmss_sales_dashboard_m.xlsx의 세일즈 대시보드를 출력한다.

m: sales_data.xlsx 파일의 버전 번호이다.
yyyymmdd_hhmmss: 세일즈 대시보드를 출력하기 위해 이 스크립트가 실행된 연월일시분초이다.

디렉토리 구조는 아래와 같다.

[세일즈 대시보드]\py
[세일즈 대시보드]\sales_data
[세일즈 대시보드]\sales_dashboard

py: 이 스크립트 파일(sales_dashboard_m.py)이 있는 디렉토리다.
sales_data: 데이터 파일 (sales_data_m.xlsx)이 있는 디렉토리다.
sales_dashboard: 세일즈 대시보드 파일들이 출력되는 디렉토리다.

이 스크립트가 실행된 후에 원본 sales_data_m.xlsx 파일은
yymmdd_hhmmss_sales_dashboard_m.xlsx 파일로 덮어써진다.
두 가지 목적이 있다.
1. sales_data_m.xlsx의 내용을 업데이트하기 위해서
2. sales_dashboard.pptx의 내용을 업데이트하기 위해서

이 스크립트가 실행되면 결국 3가지 형식의 세일즈 대시보드 파일들이 생성된다: html, xlsx, pptx.
'''


# %%
# version
# 버전 번호 정하는 방법
# Vm.n on yyyy-mm-dd:
#   m           : xlsx 템플릿의 호환성이 없어지면 바꾼다.
#   n           : 새로 배포할 때 바꾼다.
#   yyyy-mm-dd  : m.n을 바꾸지 않은 상태에서 버전 관리가 필요할 때 바꾼다.
#
# V1.0 on 2022-05-03: 첫 데모
# V1.1 on 2022-05-04: 첫 배포. 페스카로 서버에 올리고, 글로벌사업개발팀에 이-메일로 알렸다.
# V1.2 on 2022-05-06:
#   KPI 계산에 quote_price 대신 weighted_quote_price를 사용한다.
#   refactoring
#   color coding:
#       won: green, liekly: orange, quoted: red, lead: blue
#   한 SSO의 분할 payment들 중 kpi_year에 해당하는 부분만 kpi_sales에 반영되도록 하였다.
# V2.0 on 2022-05-13
#   sales와 booking을 구분하기 위해 xlsx 파일에 closed_date 컬럼을 추가했다.
#       xlsx 호환성이 없어져서 V2.0으로 버전 번호를 변경한다.
#       sales: 해당 년도에 지불이 이뤄지는 금액
#       booking: 해당 년도에 수주한 금액
#       올해 수주하여 내년에 지불이 이뤄지면 해당 금액은 booking에는 포함되나 sales에는 포함되지 않는다.
#       지난 해에 수주하여 올해 지불이 이뤄지면 해당 금액은 sales에는 포함되나 booking에는 포함되지 않는다.
#   refactoring
#   top 5, top 3 테이블을 focus_ssos 시트에 쓴다.
#   모든 작업을 마치고 기존 sales_data_m.xlsx 파일을 새로 만든 대시보드 파일로 덮어쓴다.
#       sales_dashboard.pptx를 만든다.
#       sales_dashboard.pptx에 sales_data_m.xlsx의 차트를 링크로 연결하여
#       sales_data_m.xlsx가 업데이트되면 sales_dashboard.pptx도 업데이트 되도록 한다.
#       sales_dashboard.pptx에 커멘트를 할 수 있는 슬라이드를 만든다.
#           One for analysis, one for how to take up a challenge
#   sales_data_m.xlsx의 가상 데이터를 세일즈 대시보드의 기능을 잘 데모할 수 있도록 바꾼다.
#   beautification: 그래프에 표시되는 수치의 자리수를 조정하였다.
# V2.0 on 2022-05-18
#   top 5 of snm, top 3 by solution 바 차트의 바를 sales_phase로 컬러 코딩한다.


# %%
import numpy as np
import warnings
import xlwings as xw
from pyecharts import options as opts
from pyecharts.charts import Bar, Tab, Grid, Pie
import pandas as pd
from pathlib import Path
import os
import shutil
import datetime


# %%
# constant

k_py_version = 'V2.0 released on 2022-05-18'

k_render_notebook = 0

k_company_name = 'Company Name'

k_sales_data_dir_name = 'sales_data'

k_sales_dashboard_dir_name = 'sales_dashboard'

k_base_xlsx_file_name = 'sales_data_02.xlsx'

k_sales_phases = ['won', 'likely', 'quoted',
                  'lead', 'ideation', 'lost', 'cancelled']
k_sales_phases_for_kpi = ['won', 'likely', 'quoted']

k_sales_phases_for_kpi_extended = k_sales_phases_for_kpi.copy()
k_sales_phases_for_kpi_extended.append('lead')

k_sales_phase_for_kpi_color = {
    'won': 'green',
    'likely': 'orange',
    'quoted': 'red',
}

k_sales_phase_color = k_sales_phase_for_kpi_color.copy()
k_sales_phase_color['lead'] = 'blue'
k_sales_phase_color['ideation'] = 'cyan'

k_sales_phases_for_focus_ssos = ['likely', 'quoted', 'lead', 'ideation']

k_cols_to_get_focus_ssos_for_chart = [
    'sso_n', 'sso', 'sales_phase', 'fc_amount']
k_cols_to_get_focus_ssos_for_xlsx = [
    'sso_n', 'sso', 'sales_phase', 'fc_amount', 'quote_price']

k_columns_to_rename = {
    '순번': 'entry_no',
    '고객': 'customer',
    '고객_분류': 'customer_category',
    '현황': 'status',
    '등록일': 'reg_date',
    '영업담당자': 'sales_resp',
    '기술담당자': 'rnd_resp',
    '견적일': 'quote_date',
    '견적가': 'quote_price',
    '원가': 'cost',
    '이윤': 'profit',
    '영업단계': 'sales_phase',
    '수주확률': 'probability',
    '수주확률가중견적가': 'weighted_quote_price',
    '착수일': 'proj_start_date',
    '종료일': 'proj_end_date',
    '지불방식': 'payment_terms',
    '지불액1': 'pay_amount_1',
    '지불액2': 'pay_amount_2',
    '지불액3': 'pay_amount_3',
    '지불일1': 'pay_date_1',
    '지불일2': 'pay_date_2',
    '지불일3': 'pay_date_3',
    '매출전망1': 'fc_amount_1',
    '매출전망2': 'fc_amount_2',
    '매출전망3': 'fc_amount_3',
}

k_columns_to_convert = [
    'quote_price',
    'cost',
    'profit',
    'weighted_quote_price',
    'pay_amount_1',
    'pay_amount_2',
    'pay_amount_3',
    'fc_amount_1',
    'fc_amount_2',
    'fc_amount_3',
]

k_krw_to_mkrw_factor = 0.000001


# %%
# function
def print_header(data_dir, work_dir, xlsx_file_name):
    print('\n' * 3)
    print('=' * 80)
    print()
    print(f'{k_company_name:^80}')
    print(f'Sales Dashboard Generator {k_py_version}')
    print()
    print(f'data file:')
    print(f'    {data_dir}')
    print(f'    {k_base_xlsx_file_name}')
    print(f'copy file:')
    print(f'    {work_dir}')
    print(f'    {xlsx_file_name}')
    print()
    print('=' * 80)
    print()
    return  # print_header()


def read_sales_data(xlsx_file_name):
    '''
    sales_data_m.xlsx 파일을 [file_name_prefix]_sales_data_m.xlsx로 복사한다. 이유
        1. 대시보드를 만들 때 데이터를 보관한다. 대시보드 파일 이름의 앞부분도 복사된 세일즈 데이터 파일과 동일하다. 데이터 파일과 대시보드 파일이 짝을 이룬다.
        2. sales_data_m.xlsx로 작업을 할 사람은 중단 없이 작업을 할 수 있도록 한다.
    한글 컬럼 이름을 영어로 바꾼다. 코딩의 편의를 위해서. 이렇게 할 필요가 있는가?
    일부 컬럼들의 타입을 변경한다.
    '''
    with warnings.catch_warnings():
        # sales_data_m.xlsx 파일에 입력값을 선택하는 리스트가 있다.
        # pandas에는 리스트처럼 입력값의 타당성을 검증하는 기능이 없다.
        # 관련 경고를 출력하지 않다록 한다.
        warnings.simplefilter("ignore")

        df = pd.read_excel(xlsx_file_name, sheet_name='data', skiprows=3)

    df.rename(columns=k_columns_to_rename, inplace=True)

    # 엑셀 파일의 컬럼 데이터 형식이 pandas에서 필요한 형식과 다르다.
    df['probability'] = df['probability'].astype('category')

    df['sc_2'].fillna(0, inplace=True)
    df['sc_2'] = df['sc_2'].astype('int64')

    df['cost'].fillna(0, inplace=True)
    df['cost'] = df['cost'].astype('int64')

    return df  # read_sales_data()


def flatten_per_pay_date(df):
    '''
    sales_data_m.xlsx 파일의 1행을 pay_date_1, pay_date_2, pay_date_3에 따라
    3행으로 펼친다(flatten).
    펼칠 때, pay_date_1, pay_date_2, pay_date_3 컬럼들을 pay_date 컬럼으로,
    fc_amount_1, fc_amount_2, fc_amount_3 컬럼들을 fc_amount 컬럼으로 만든다.
    '''
    dfs_to_concat = []
    for i in range(1, 4):
        df_temp = df.copy()
        df_temp['pay_date'] = df_temp[f'pay_date_{i}']
        df_temp['fc_amount'] = df_temp[f'fc_amount_{i}']
        # df_temp.drop(
        #     ['pay_date_1', 'pay_date_2', 'pay_date_3', 'fc_amount_1', 'fc_amount_2', 'fc_amount_3'], axis=1, inplace=True)

        dfs_to_concat.append(df_temp)

    df_flat = pd.concat(dfs_to_concat)
    df_temp.drop(['pay_date_1', 'pay_date_2', 'pay_date_3',
                  'fc_amount_1', 'fc_amount_2', 'fc_amount_3'], axis=1, inplace=True)
    df_flat['pay_year'] = df_flat['pay_date'].dt.year
    df_flat['pay_month'] = df_flat['pay_date'].dt.month
    df_flat['closed_year'] = df_flat['closed_date'].dt.year
    df_flat['closed_month'] = df_flat['closed_date'].dt.year

    # sort on pay_date and sales_phase
    df_flat.sort_values(
        by=['pay_date', 'sales_phase'],
        ascending=[True, False],
        axis=0,
        inplace=True)

    # df_flat.set_index('pay_date', inplace=True)

    return df_flat    # concat_fc_data()


def write_sales_fc(wb, sheet_name, titles, df_rc, df):
    '''
    write data to xlsx
        wb          : workbook. The xlsx to write data to.
        sheet_name  : worksheet.
        title_rc    : row and column of the cell where the title will be wrtten
        title       : title
        df_rc       : row and column of the cell where the df will be written
        df          : dataframe containing the data to write
    '''
    ws = wb.sheets(sheet_name)
    for rc, title in titles.items():
        ws.range(rc).value = title
    ws.range(df_rc).value = df
    return  # write_sales_fc()


def draw_kpi_bar_charts(df_kpi):
    # KPI bar chart
    # html 대시보드에 그리지 않는다.
    bar_chart_kpi_total = Bar()
    bar_chart_kpi_total.add_xaxis(df_kpi['sales_phase'].to_list())
    bar_chart_kpi_total.add_yaxis('매출', df_kpi['total'].to_list())
    bar_chart_kpi_total.set_global_opts(
        title_opts=opts.TitleOpts(title=f'{kpi_year}년 매출 전망 [MKRW]'))

    bar_chart_kpi_count = Bar()
    bar_chart_kpi_count.add_xaxis(df_kpi['sales_phase'].to_list())
    bar_chart_kpi_count.add_yaxis('건', df_kpi['count'].to_list())
    bar_chart_kpi_count.set_global_opts(
        title_opts=opts.TitleOpts(title=f'{kpi_year}년 SSO [건]'))

    bar_chart_kpi_average = Bar()
    bar_chart_kpi_average.add_xaxis(df_kpi['sales_phase'].to_list())
    bar_chart_kpi_average.add_yaxis('매출/건', df_kpi['average'].to_list())
    bar_chart_kpi_average.set_global_opts(
        title_opts=opts.TitleOpts(title=f'{kpi_year}년 매출/건 [MKRW/건]'))

    # bar_chart_sales_fc_by_month.set_series_opts(label_opts=opts.LabelOpts(position='inside'))
    # bar_chart_sales_fc_by_month.set_global_opts(title_opts=opts.TitleOpts(title=f'{year}년 월별 누적 매출 전망'),
    #     legend_opts=opts.LegendOpts(pos_top='middle', pos_right='right')
    #     )

    grid = Grid()
    # pos_bottom은 아래쪽 여백을 의미한다.
    grid.add(bar_chart_kpi_total, grid_opts=opts.GridOpts(pos_bottom='70%'))
    grid.add(bar_chart_kpi_count, grid_opts=opts.GridOpts(
        pos_top='40%', pos_bottom='40%'))
    grid.add(bar_chart_kpi_average, grid_opts=opts.GridOpts(pos_top='70%'))
    return grid  # draw_kpi_bar_charts


def draw_kpi_pie_charts(df_kpi, title, subtitle):
    '''
    KPI pie chart
    '''

    k_pie_n = 3

    interval = int(100 / k_pie_n)   # 100은 차트가 그려질 캔버스 폭의 100%를 의미하는구나.
    offset = interval // 2

    pie_kpi = Pie(init_opts=opts.InitOpts(width='1200px', height='600px'))
    kpi_names = ['total', 'count', 'average']
    titles = ['매출(wlq) [MKRW]', 'SSO [건]', '매출(wlq)/SSO [MKRW/건]']
    max_radius = max(offset, 10)
    min_radius = max(max_radius - 10, 5)
    for i in range(k_pie_n):
        center_x = i * interval + offset

        pie_kpi.add(
            titles[i],
            [list(z) for z in zip(df_kpi['sales_phase'].to_list(),
                                  df_kpi[kpi_names[i]].round(1).to_list())],
            radius=[f'{min_radius}%', f'{max_radius}%'],
            center=[f'{center_x}%', '50%'],
            rosetype="radius",
        )
        pie_kpi.set_colors(list(k_sales_phase_for_kpi_color.values()))

    # Custom data labels
    pie_kpi.set_global_opts(
        title_opts=opts.TitleOpts(
            title=title,
            subtitle=subtitle),
        legend_opts=opts.LegendOpts(pos_top='bottom', pos_left='center')
    )

    pie_kpi.set_series_opts(
        label_opts=opts.LabelOpts(
            # position='top',
            # color='red',
            # font_family='Arial',
            # font_size=12,
            # font_style='italic',
            # interval=1,
            # formatter='{b}:{d}%'
            formatter='{b}: {c}'
        )
    )
    return pie_kpi


def draw_bar_chart_sales_fc(df, title, subtitle):
    # bar chart
    bar_chart = Bar()

    bar_chart.add_xaxis(list(df.index))

    for sales_phase in k_sales_phases_for_kpi:
        bar_chart.add_yaxis(
            sales_phase,
            df[sales_phase].round(0).tolist(),
            stack=True,
            itemstyle_opts=opts.ItemStyleOpts(
                color=k_sales_phase_color[sales_phase]
            )
        )
    bar_chart.set_series_opts(
        label_opts=opts.LabelOpts(position='inside'))

    label_angle = 0
    if len(df.index) > 5:
        label_angle = -90

    bar_chart.set_global_opts(
        title_opts=opts.TitleOpts(
            title=title,
            subtitle=subtitle
        ),
        legend_opts=opts.LegendOpts(
            pos_top='middle',
            pos_right='right'
        ),
        xaxis_opts=opts.AxisOpts(
            axislabel_opts=opts.LabelOpts(
                rotate=label_angle,
                margin=8
            )
        ),
    )
    return bar_chart  # draw_bar_chart_sales_fc


def draw_bar_chart_top_sso(df, title, subtitle):

    bar_chart = Bar()
    ssos = list(df['sso'].unique())
    print(f'{ssos = }')
    ssos = ssos[::-1]
    bar_chart.add_xaxis(ssos)

    ssos_len = len(ssos)
    for sales_phase in k_sales_phases_for_focus_ssos:
        if sales_phase in df.columns:
            fc_amounts = df[sales_phase].apply(np.ceil).to_list()
        else:
            fc_amounts = [0.0] * ssos_len
        fc_amounts = fc_amounts[::-1]
        print(f'{sales_phase:12} {fc_amounts}')
        bar_chart.add_yaxis(
            sales_phase,
            fc_amounts,
            stack=True,
            itemstyle_opts=opts.ItemStyleOpts(
                color=k_sales_phase_color[sales_phase])
        )

    bar_chart.reversal_axis()
    bar_chart.set_series_opts(label_opts=opts.LabelOpts(position='inside'))
    bar_chart.set_global_opts(
        title_opts=opts.TitleOpts(title=title, subtitle=subtitle),
        legend_opts=opts.LegendOpts(pos_top='bottom', pos_right='center'),
    )

    # bar_chart = Bar()

    # # ssos = df_temp.index.to_list()
    # ssos = df['sso'].to_list()
    # ssos = ssos[::-1]
    # bar_chart.add_xaxis(ssos)

    # # for sales_phase in k_sales_phases_for_kpi_extended:
    # potentials = df['fc_amount'].apply(np.ceil).to_list()
    # potentials = potentials[::-1]
    # bar_chart.add_yaxis('', potentials)

    # bar_chart.reversal_axis()
    # bar_chart.set_series_opts(
    #     label_opts=opts.LabelOpts(position='right'))
    # bar_chart.set_global_opts(
    #     title_opts=opts.TitleOpts(title=title, subtitle=subtitle),
    #     legend_opts=opts.LegendOpts(pos_top='bottom', pos_right='center'),
    # )
    return bar_chart  # draw_bar_chart_top_sso


def main():
    return      # main()


# main
# TODO
# if __name__ == '__main__':
#     main()

# TODO
# 직전 연도 sales, booking를 올해의 세일즈, 부킹과 비교한다.
#   sales_data.xlsx에 있는 연도들을 gui checkbox에 리스트하고,
#   선택된 연도들의 그래프를 그린다.
#   지난 해는 선 그래프 (sales_phases가 won인 수치만 표시되니까.), 올해는 막대 그래프

# TODO
# pyecharts의 tab 별로 흩어져있는 html 출력들을 한 곳으로 모으고, xlsx 출력들도 한 곳으로 모은다.
# 계산, html 출력, xlsx 출력 순서로 실행되도록 구조를 바꿔 코드를 쉽게 이해하고 수정할 수 있도록 한다.
# 이렇게 구조를 바꾸는 것이 더 코드를 번거롭지 않을까?

# ----- 매출 전망을 계산하기 전 준비 작업 -----
# sales_data_m.xlsx의 복사본으로 작업한다.

# %%
# paths on PC for development
home_dir = Path.cwd()    # the directory where this python script file is

# paths on the server for use
# home_dir = Path('\\\\192.168.0.200\\snm$\\0200.사업개발\\sales_dashboard')

data_dir = home_dir.parent / k_sales_data_dir_name
work_dir = home_dir.parent / k_sales_dashboard_dir_name

print(f'home directory: {home_dir}')
print(f'date directory: {data_dir}')
print(f'work directory: {work_dir}')

file_name_prefix = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
xlsx_file_name = file_name_prefix + '_' + k_base_xlsx_file_name

# error handling: sales_data_m.xlsx is open already (by Excel)
os.chdir(work_dir)
while True:
    try:
        shutil.copyfile(data_dir/k_base_xlsx_file_name,
                        work_dir/xlsx_file_name)
        break
    except:
        print('The sales data file may be open.')
        input('Press Enter when ready to proceed.')

print_header(data_dir, work_dir, xlsx_file_name)


# %%
df = read_sales_data(xlsx_file_name)


# %%
# 원 단위로 표시하니 차트가 너무 복잡하다. 백만원 단위로 변경한다.
df[k_columns_to_convert] = df[k_columns_to_convert] * \
    k_krw_to_mkrw_factor


# %%
# 엑셀 파일의 한 행에는 pay_date_1, pay_date_2, pay_date_3의 컬럼이 있다.
# pay_date_1, pay_date_2, pay_date_3를 기준으로 한 행을 세 행으로 펼친다.
df_flat = flatten_per_pay_date(df)


# %%
# TODO
# 계산할 연월을 선택할 수 있도록 한다.
# 디폴트는 오늘이 속한 연월이다.
today = datetime.date.today()
kpi_year = today.year
kpi_month = today.month


# %%
# sales_data_m.xlsx의 대시보드 데이터를 업데이트하기 위해 sales_data_m.xlsx를 wb로 연다.
# [important] read_sales_data() 후에 실행해야 한다. 아니면 file access 에러가 발생한다.
app = xw.App(visible=False)
wb = xw.Book(xlsx_file_name)


# %% [markdown]
# KPI by sales


# %%
# sales forecast amount per sales phase
kpi_sales_total = df_flat.query('pay_year==@kpi_year')[['sales_phase', 'fc_amount']].groupby(
    'sales_phase')['fc_amount'].sum()[k_sales_phases_for_kpi]
# print(kpi_sales_total)


# %%
kpi_sales_total_sum = kpi_sales_total.sum()
# print(f'{kpi_sales_total_sum = }')


# %%
# sales sso count per sales phase
df_kpi_sales_sso = (df_flat.query('pay_year==@kpi_year')[['sso_n', 'sales_phase']].
                    groupby(by=['sso_n', 'sales_phase'])['sso_n'].
                    count().unstack()).fillna(0)[k_sales_phases_for_kpi]
# payment가 pay_date_1, pay_date_2, pay_date_3로 분리되므로 동일 sso에 pay_year가 복수로 있을 수 있다.
df_kpi_sales_sso[df_kpi_sales_sso > 0] = 1
df_kpi_sales_sso['sum'] = df_kpi_sales_sso.sum(axis='columns')

# 동일 sso에 여러 position이 포함되어 있을 수 있다.
# position별로 sales_phase가 다를 수 있다.
# 예)
# 한 sso 견적서에 제품 포지션, 엔지니어링 포지션, 컨설팅 포지션이 있다.
# 제품은 구매 확률이 높고, 엔지니어링은 취소되었다면,
# 제품, 엔지니어링, 컨설팅의 sales_phase는 각각 likely, cancelled, quoted로 분리하여 관리될 수 있다.
# sso는 하나이지만 제품, 엔지니어링, 컨설팅을 각각 1/3개 sso로 계산한다.

# divide by 0를 방지하기 위해서 'sum'이 0인 행은 제외한다.
df_kpi_sales_sso = df_kpi_sales_sso[df_kpi_sales_sso['sum'] > 0]
for column in k_sales_phases_for_kpi:
    df_kpi_sales_sso[column] = df_kpi_sales_sso[column] / \
        df_kpi_sales_sso['sum']

df_kpi_sales_sso.drop(columns=['sum'], inplace=True)
kpi_sales_sso_count = df_kpi_sales_sso.sum(axis='rows')
kpi_sales_sso_count.name = 'count'
# print(f'{kpi_sales_sso_count = }')


# %%
kpi_sales_sso_count_sum = kpi_sales_sso_count.sum()
# print(f'{kpi_sales_sso_count_sum = }')


# %%
# average forecast sales (= sales forecast amount / sales sso count) per sales phase
if kpi_sales_sso_count_sum:
    kpi_sales_average_sum = 0
else:
    kpi_sales_average_sum = kpi_sales_total_sum / kpi_sales_sso_count_sum


# %%
# make df_sales_kpi
df_sales_kpi = pd.concat([kpi_sales_total, kpi_sales_sso_count],
                         axis='columns').reset_index()
df_sales_kpi['average'] = df_sales_kpi['fc_amount'] / df_sales_kpi['count']
df_sales_kpi.rename(columns={'fc_amount': 'total'}, inplace=True)
# print(f'{df_sales_kpi = }')


# %%
# kpi sales chart
title = f'{kpi_year}년 사업 개발 KPI - Sales'
subtitle = f'매출(wlq) {kpi_sales_total_sum:>15,.1f} MKRW (좌)\n\nSSO {kpi_sales_sso_count_sum:>26,} 건\n\n건당 매출 {kpi_sales_average_sum:>15,.1f} MKRW/건 (우)'
kpi_sales_pie_chart = draw_kpi_pie_charts(df_sales_kpi, title, subtitle)


# %%
if k_render_notebook:
    kpi_sales_pie_chart.render_notebook()


# %%
# kpi sales to xlsx
write_sales_fc(wb, 'kpi_sales',
               {
                   'C1': f'매출 전망\n{kpi_year}년 {kpi_month}월\n{kpi_sales_total_sum:,.0f} MKRW',
                   'D1': f'SSO\n{kpi_year}년 {kpi_month}월\n{kpi_sales_sso_count_sum:,} 건',
                   'E1': f'SSO당 매출 전망 평균\n{kpi_year}년 {kpi_month}월\n{kpi_sales_average_sum:,.0f} MKRW/건'
               },
               'A3', df_sales_kpi)


# %% [markdown]
# KPI by booking
'''
sales는 매출 발생 연도가 당해 연도인 매출이다. 실제 수주는 이전 연도가 될 수도 있다.
booking은 수주가 당해 연도인 SSO의 총 매출이다. 실제 매출은 다음 연도에 이후에 발생할 수도 있다.
'''


# %%
# booking amount per sales phase

# TODO
# kpi_sales와 kpi_booking이 매우 비슷하다. 한 함수로 만들 수 없을까?
# sales는 pay_year==@kpi_year인데 반해 booking은 closed_year==@kpi_year이다.

kpi_booking_total = df_flat.query('closed_year==@kpi_year')[['sales_phase', 'fc_amount']].groupby(
    'sales_phase')['fc_amount'].sum()[k_sales_phases_for_kpi]
# print(f'{kpi_booking_total = }')


# %%
kpi_booking_total_sum = kpi_booking_total.sum()
# print(f'{kpi_booking_total_sum = }')


# %%
# booking sso count per sales phase

# TODO
# sso_count 부분을 함수로 만들고 sales와 booking에 각각 적용하면 되겠다.

df_kpi_booking_sso = (df_flat.query('closed_year==@kpi_year')[['sso_n', 'sales_phase']].
                      groupby(by=['sso_n', 'sales_phase'])['sso_n'].
                      count().unstack()).fillna(0)[k_sales_phases_for_kpi]
df_kpi_booking_sso[df_kpi_booking_sso > 0] = 1
df_kpi_booking_sso['sum'] = df_kpi_booking_sso.sum(axis='columns')
df_kpi_booking_sso = df_kpi_booking_sso[df_kpi_booking_sso['sum'] > 0]
for column in k_sales_phases_for_kpi:
    df_kpi_booking_sso[column] = df_kpi_booking_sso[column] / \
        df_kpi_booking_sso['sum']
df_kpi_booking_sso.drop(columns=['sum'], inplace=True)
kpi_booking_sso_count = df_kpi_booking_sso.sum(axis='rows')
kpi_booking_sso_count.name = 'count'
# print(f'{kpi_booking_sso_count = }')


# %%
kpi_booking_sso_count_sum = kpi_booking_sso_count.sum()
# print(f'{kpi_booking_sso_count_sum = }')


# %%
if kpi_booking_sso_count_sum:
    kpi_booking_average_sum = 0
else:
    kpi_booking_average_sum = kpi_booking_total_sum / kpi_booking_sso_count_sum


# %%
# make df_sales_kpi
df_booking_kpi = pd.concat([kpi_booking_total, kpi_booking_sso_count],
                           axis='columns').reset_index()
df_booking_kpi['average'] = df_booking_kpi['fc_amount'] / \
    df_booking_kpi['count']
df_booking_kpi.rename(columns={'fc_amount': 'total'}, inplace=True)
# print(f'{df_booking_kpi = }')


# %%
# kpi booking chart
# kpi_bar_chart_grid = draw_kpi_bar_charts(df_sales)
title = f'{kpi_year}년 사업 개발 KPI - Booking'
subtitle = f'매출(wlq) {kpi_booking_total_sum:>15,.1f} MKRW (좌)\n\nSSO {kpi_booking_sso_count_sum:>26,} 건\n\n건당 매출 {kpi_booking_average_sum:>15,.1f} MKRW/건 (우)'
kpi_booking_pie_chart = draw_kpi_pie_charts(
    df_booking_kpi, title, subtitle)


# %%
if k_render_notebook:
    kpi_booking_pie_chart.render_notebook()


# %%
# kpi bookign to xlsx
if 'kpi_booking' not in [sheet.name for sheet in wb.sheets]:
    wb.sheets.add('kpi_booking')

write_sales_fc(wb, 'kpi_booking',
               {
                   'C1': f'수주 전망\n{kpi_year}년 {kpi_month}월\n{kpi_sales_total_sum:,.0f} MKRW',
                   'D1': f'SSO\n{kpi_year}년 {kpi_month}월\n{kpi_booking_sso_count_sum:,} 건',
                   'E1': f'SSO당 수주 전망 평균\n{kpi_year}년 {kpi_month}월\n{kpi_sales_average_sum:,.0f} MKRW/건'
               },
               'A3', df_booking_kpi)


# %% [markdown]
# sales forecast by month


# %%
df_selection = df_flat.query('pay_year == @kpi_year')
df_sales_monthly_sum = df_selection.groupby(
    by=[pd.Grouper(key='pay_date', freq='M'), 'sales_phase'])[
    'fc_amount'].sum().reset_index().query('sales_phase in @k_sales_phases_for_kpi')
df_sales_monthly_sum.sort_values(by=['pay_date', 'sales_phase'],
                                 ascending=[True, False], inplace=True)
df_sales_monthly_sum_by_sales_phase = df_sales_monthly_sum.groupby(
    [pd.Grouper(key='pay_date', freq='M'), 'sales_phase']).sum().unstack().fillna(0)
df_sales_monthly_sum_by_sales_phase.columns = ['likely', 'quoted', 'won']

cumsums = []
for column in df_sales_monthly_sum_by_sales_phase.columns:
    cumsums.append(df_sales_monthly_sum_by_sales_phase[column].cumsum())
df_sales_monthly_cumsum_by_sales_phase = pd.concat(
    cumsums, axis='columns').reset_index()


# %%
# bar chart
bar_chart_sales_fc_by_month = Bar()
months = list(range(1, 13))
bar_chart_sales_fc_by_month.add_xaxis(months)

fc_dict = dict()
df_sales_monthly_cumsum_by_sales_phase['pay_month'] = df_sales_monthly_cumsum_by_sales_phase['pay_date'].dt.month
for sales_phase in k_sales_phases_for_kpi:
    df_temp = df_sales_monthly_cumsum_by_sales_phase[[
        'pay_month', sales_phase]]
    pay_months = list(df_temp['pay_month'])
    cumsums = df_temp[sales_phase].round(0).to_list()

    for month in months:
        if month not in pay_months:
            pay_months.insert(month - 1, month)
            if month == 1:
                cumsums.insert(month - 1, 0.0)
            else:
                prev_month_cumsum = cumsums[month - 2]
                cumsums.insert(month - 1, prev_month_cumsum)

    fc_dict[sales_phase] = cumsums

    bar_chart_sales_fc_by_month.add_yaxis(
        sales_phase,
        cumsums,
        stack=True,
        itemstyle_opts=opts.ItemStyleOpts(
            color=k_sales_phase_color[sales_phase]
        )
    )

bar_chart_sales_fc_by_month.set_series_opts(
    label_opts=opts.LabelOpts(position='inside'))
bar_chart_sales_fc_by_month.set_global_opts(title_opts=opts.TitleOpts(
    title=f'{kpi_year}년 월별 누적 매출 전망',
    subtitle=f'매출(wlq) {kpi_sales_total_sum:>15,.1f} MKRW'),
    legend_opts=opts.LegendOpts(pos_top='middle', pos_right='right')
)


# %%
if k_render_notebook:
    bar_chart_sales_fc_by_month.render_notebook()
# bar_chart_sales_fc_by_month.render(f'{file_name_prefix}_bar_fc_by_month.html')


# %%
# sales forecast by month to xlsx
df_fc_monthly_cumsum_to_plot = pd.DataFrame(fc_dict)

if 'fc_by_month' not in [sheet.name for sheet in wb.sheets]:
    wb.sheets.add('fc_by_month')
write_sales_fc(wb, 'fc_by_month',
               {'A1': f'매출 전망\n{kpi_year}년 {kpi_month}월\n{kpi_sales_total_sum:,.0f} MKRW'},
               'A3', df_fc_monthly_cumsum_to_plot[k_sales_phases_for_kpi])
ws = wb.sheets('fc_by_month')
ws.range('A3').value = '월'
ws.range('A4').options(transpose=True).value = list(range(1, 13))


# %% [markdown]
# sales forecast by solution field


# %%
df_selection = df_flat.query('pay_year==@kpi_year')
df_sales_fc_by_sf = df_selection.groupby(['solution_field', 'sales_phase'])[
    'fc_amount'].sum().unstack().fillna(0)


# %%
# sales forecast by solution field bar chart
bar_chart_sales_fc_by_sf = draw_bar_chart_sales_fc(
    df_sales_fc_by_sf,
    f'{kpi_year}년 솔루션 필드별 매출 전망',
    f'매출(wlq) {kpi_sales_total_sum:>15,.1f} MKRW'
)

# %%
if k_render_notebook:
    bar_chart_sales_fc_by_sf.render_notebook()


# %%
# sales forecast by solution field to xlsx
if 'fc_by_sf' not in [sheet.name for sheet in wb.sheets]:
    wb.sheets.add('fc_by_sf')
write_sales_fc(wb, 'fc_by_sf',
               {'A1': f'솔루션 필드별 매출 전망\n{kpi_year}년 {kpi_month}월\n{kpi_sales_total_sum:,.0f} MKRW'},
               'A3', df_sales_fc_by_sf[k_sales_phases_for_kpi_extended])


# %% [markdown]
# sales forecast by solution


# %%
df_selection = df_flat.query('pay_year==@kpi_year')
df_sales_fc_by_solution = df_selection.groupby(['solution', 'sales_phase'])[
    'fc_amount'].sum().unstack().fillna(0)
df_sales_fc_by_solution.index = df_sales_fc_by_solution.index.str.strip()


# %%
# sales forecast by solution bar chart
bar_chart_sales_fc_by_solution = draw_bar_chart_sales_fc(
    df_sales_fc_by_solution,
    f'{kpi_year}년 솔루션별 매출 전망',
    f'매출(wlq) {kpi_sales_total_sum:>15,.1f} MKRW'
)


# %%
if k_render_notebook:
    bar_chart_sales_fc_by_solution.render_notebook()


# %%
# sales forecast by solution to xlsx
if 'fc_by_solution' not in [sheet.name for sheet in wb.sheets]:
    wb.sheets.add('fc_by_solution')
write_sales_fc(wb, 'fc_by_solution',
               {'A1': f'솔루션별 매출 전망\n{kpi_year}년 {kpi_month}월\n{kpi_sales_total_sum:,.0f} MKRW'},
               'A3', df_sales_fc_by_solution[k_sales_phases_for_kpi_extended])


# %% [markdown]
# sales forecast by customer
df_selection = df_flat.query('pay_year==@kpi_year')
df_sales_fc_by_customer = df_selection.groupby(['customer', 'sales_phase'])[
    'fc_amount'].sum().unstack()


# %%
# sales forecast by customer bar chart
bar_chart_sales_fc_by_customer = draw_bar_chart_sales_fc(
    df_sales_fc_by_customer,
    f'{kpi_year}년 고객사별 매출 전망',
    f'매출(wlq) {kpi_sales_total_sum:>15,.1f} MKRW'
)


# %%
if k_render_notebook:
    bar_chart_sales_fc_by_customer.render_notebook()


# %%
# sales forecast by customer to xlsx
if 'fc_by_customer' not in [sheet.name for sheet in wb.sheets]:
    wb.sheets.add('fc_by_customer')
write_sales_fc(wb, 'fc_by_customer',
               {'A1': f'고객사별 매출 전망\n{kpi_year}년 {kpi_month}월\n{kpi_sales_total_sum:,.0f} MKRW'},
               'A3', df_sales_fc_by_customer[k_sales_phases_for_kpi_extended])


# %% [markdown]
# sales forecast by customer category -----


# %%
# sales forecast by customer category

df_selection = df_flat.query('pay_year==@kpi_year')
df_sales_fc_by_customer_category = df_selection.groupby(['customer_category', 'sales_phase'])[
    'fc_amount'].sum().unstack().fillna(0)


# %%
# sales forecast by customer category bar chart
bar_chart_sales_fc_by_customer_category = draw_bar_chart_sales_fc(
    df_sales_fc_by_customer_category,
    f'{kpi_year}년 고객군별 매출 전망',
    f'매출(wlq) {kpi_sales_total_sum:>15,.1f} MKRW'
)


# %%
if k_render_notebook:
    bar_chart_sales_fc_by_customer_category.render_notebook()


# %%
# sales forecast by customer category to xlsx
if 'cust_cat' not in [sheet.name for sheet in wb.sheets]:
    wb.sheets.add('cust_cat')

write_sales_fc(wb, 'cust_cat',
               {'A1': f'고객군별 매출 전망\n{kpi_year}년 {kpi_month}월\n{kpi_sales_total_sum:,.0f} MKRW'},
               'A3', df_sales_fc_by_customer_category)


# %%
# TODO
# ----- lead -----
# 솔루션 탭 바로 다음에 lead 탭을 만든다.
# likely, quoted, lead, ideation를 대상으로 kpi 차트와 같은 유형의 시각화를 한다.


# %% [markdown]
# top 5 ssos of snm(Sales & Marketing)


# %%
sso_ns = df_flat.query('pay_year==@kpi_year and sales_phase in @k_sales_phases_for_focus_ssos')[
    k_cols_to_get_focus_ssos_for_chart].groupby('sso_n').sum().nlargest(5, 'fc_amount').index.to_list()

df_top_5_ssos_of_snm = df_flat.query('pay_year==@kpi_year and sso_n in @sso_ns')[
    k_cols_to_get_focus_ssos_for_chart].groupby(['sso_n', 'sso', 'sales_phase']).sum().unstack().fillna(0)
df_top_5_ssos_of_snm.reset_index(inplace=True)

# rename columns to convert multi-index to single index
new_columns = []
for col in df_top_5_ssos_of_snm.columns:
    if col[1]:
        new_columns.append(col[1])
    else:
        new_columns.append(col[0])
df_top_5_ssos_of_snm.columns = new_columns

# sort per fc_amount
df_top_5_ssos_of_snm['sorter'] = df_top_5_ssos_of_snm['sso_n'].map(
    lambda x: sso_ns.index(x))
df_top_5_ssos_of_snm.sort_values('sorter', inplace=True)


# %%
# top 5 ssos of snm bar chart
title = f'사업개발팀 Top 5 SSO - {kpi_year}년 {kpi_month}월'
subtitle = f'Potential: 예상 매출 x 수주 가능성 [MKRW]'
bar_chart_top_5_ssos_of_snm = draw_bar_chart_top_sso(
    df_top_5_ssos_of_snm, title, subtitle)


# %%
if k_render_notebook:
    bar_chart_top_5_ssos_of_snm.render_notebook()


# %%
# top 5 ssos of snm to xlsx
if 'focus_ssos' not in [sheet.name for sheet in wb.sheets]:
    wb.sheets.add('focus_ssos')
else:
    wb.sheets['focus_ssos'].clear()

k_cols_to_get_focus_ssos_for_xlsx = [
    'sso_n', 'sso', 'sales_phase', 'fc_amount', 'quote_price']
df_top_5_ssos_of_snm_for_xlsx = df_flat.query('pay_year==@kpi_year and sales_phase in @k_sales_phases_for_focus_ssos')[
    k_cols_to_get_focus_ssos_for_xlsx].groupby(['sso_n', 'sso'], as_index=False).sum().nlargest(5, 'fc_amount')
write_sales_fc(wb, 'focus_ssos',
               {'A1': title},
               'A3', df_top_5_ssos_of_snm_for_xlsx)


# %% [markdown]
# Top 3 ssos by solution


# %%
# TODO
# solution 별로 차트를 그린다.
# 한 페이지에 그리는 방법은 무엇인가?

df_flat['solution'] = df_flat['solution'].str.strip()   # solution의 \n을 제거한다.
solutions = list(df_flat['solution'].unique())

# print()
# print(f'{solutions = }')
# print()

dfs_to_concat = []
for solution in solutions:
    df_temp = df_flat\
        .query('(pay_year == @kpi_year) and (solution == @solution) and (sales_phase in @k_sales_phases_for_focus_ssos)')[['sso_n', 'sso', 'solution', 'sales_phase', 'fc_amount', 'quote_price']]\
        .groupby(['solution', 'sso_n', 'sso', 'sales_phase'], as_index=False).sum()\
        .nlargest(5, 'fc_amount')

    # print(f'{solution = }')
    # print(df_temp)
    # print()

    if not df_temp.empty:
        dfs_to_concat.append(df_temp)

if dfs_to_concat:
    df_top_3_ssos_by_solution = pd.concat(
        dfs_to_concat, ignore_index=True)

df_top_3_ssos_by_solution_for_chart = df_flat.query('pay_year==@kpi_year and sales_phase in @k_sales_phases_for_focus_ssos')[
    'solution sso_n sso sales_phase fc_amount'.split()].groupby(['solution', 'sso_n', 'sso', 'sales_phase']).sum().unstack().fillna(0).reset_index()

df_top_3_ssos_by_solution_for_chart.sso = df_top_3_ssos_by_solution_for_chart.sso + ' ' + \
    df_top_3_ssos_by_solution_for_chart.solution.str[:2]
new_columns = []
for col in df_top_3_ssos_by_solution_for_chart.columns:
    if col[1]:
        new_columns.append(col[1])
    else:
        new_columns.append(col[0])
df_top_3_ssos_by_solution_for_chart.columns = new_columns


# %%
# Top 3 ssos by solution bar chart
title = f'솔루션별 Top 3 SSO - {kpi_year}년 {kpi_month}월'
subtitle = f'Potential: 예상 매출 x 수주 가능성 [MKRW]'
bar_chart_top_3_ssos_by_solution = draw_bar_chart_top_sso(
    df_top_3_ssos_by_solution_for_chart, title, subtitle)


# %%
if k_render_notebook:
    bar_chart_top_3_ssos_by_solution.render_notebook()


# %%
# Top 3 ssos by solution to xlsx
write_sales_fc(wb, 'focus_ssos',
               {'A11': title},
               'A13', df_top_3_ssos_by_solution)


# %%
# write to html
tab = Tab(page_title='sales dashboard')
tab.add(kpi_sales_pie_chart, 'sales')
tab.add(bar_chart_sales_fc_by_month, 'month')
tab.add(bar_chart_sales_fc_by_sf, 'solution field')
tab.add(bar_chart_sales_fc_by_solution, 'solution')
tab.add(bar_chart_sales_fc_by_customer, 'customer')
tab.add(bar_chart_sales_fc_by_customer_category, 'cust cat')
tab.add(bar_chart_top_5_ssos_of_snm, 'top 5 of snm')
tab.add(bar_chart_top_3_ssos_by_solution, 'top 3 by solution')
tab.add(kpi_booking_pie_chart, 'booking')
sales_dashboard_html_file_name = Path.cwd(
) / f'{file_name_prefix}_sales_dashboard.html'
tab.render(sales_dashboard_html_file_name)
os.startfile(sales_dashboard_html_file_name)


# %%
wb.save()
wb.close()


# %%
print('\n' * 3)
print(f'overwrite {data_dir / k_base_xlsx_file_name}')
print(f'with {xlsx_file_name}')
shutil.copyfile(xlsx_file_name, data_dir / k_base_xlsx_file_name)
os.startfile(xlsx_file_name)
os.chdir(home_dir)

print()
print('sales dashboard generation completed successfully')
print()
