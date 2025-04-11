import streamlit as st
import pandas as pd
from datetime import datetime, date
import pandas.tseries.offsets as offsets
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import openpyxl

# 페이지 설정
st.set_page_config(page_title="자금일보", layout="wide")

# 파일 업로드 (사이드바)
uploaded_file = st.sidebar.file_uploader("엑셀 파일을 업로드하세요", type=['xlsx', 'xlsm', 'xls'], accept_multiple_files=False)

if uploaded_file is not None:
    # Daily 시트 로드
    df_daily = pd.read_excel(uploaded_file, sheet_name="Daily")
    df_daily['지출일'] = pd.to_datetime(df_daily['지출일'])
    
    # 탭 생성
    tab1, tab2 = st.tabs(["💰 자금일보", "💸 현금흐름표"])
    
    # 첫 번째 탭: 자금일보
    with tab1:
        # 날짜 입력
        col1, col2 = st.columns(2)  # 3개에서 2개로 변경
        with col1:
            start_date = st.date_input("시작일")
            start_datetime = pd.to_datetime(start_date)
        with col2:
            end_date = st.date_input("종료일")
            end_datetime = pd.to_datetime(end_date)
            end_datetime = end_datetime + offsets.Day(1) - offsets.Second(1)

        # 1. 기준기간 주요 현황
        st.header("1. 기준기간 주요 현황")
        
        # 기말잔액 계산
        final_balance = df_daily[df_daily['지출일'] <= end_datetime]['집행 금액'].sum()
        
        # 입금액 계산 ("계좌 대체" 제외)
        total_deposit = df_daily[
            (df_daily['지출일'] >= start_datetime) & 
            (df_daily['지출일'] <= end_datetime) &
            (df_daily['입금'].notna()) &
            (df_daily['현금흐름 대분류'] != "계좌 대체")  # 계좌 대체 제외
        ]['입금'].sum()
        
        # 출금액 계산 ("계좌 대체" 제외)
        total_withdrawal = -1 * df_daily[
            (df_daily['지출일'] >= start_datetime) & 
            (df_daily['지출일'] <= end_datetime) &
            (df_daily['출금'].notna()) &
            (df_daily['현금흐름 대분류'] != "계좌 대체")  # 계좌 대체 제외
        ]['출금'].sum()

        # CSS 스타일 정의
        st.markdown("""
            <style>
            .metric-container {
                display: flex;
                align-items: center;
                gap: 10px;
            }
            .amount-box {
                background-color: #f0f2f6;
                border-radius: 10px;
                padding: 15px 20px;
                flex-grow: 1;
            }
            .amount-value {
                font-size: 1.5rem;
                font-weight: bold;
                color: #0F1116;
                text-align: right;
            }
            .amount-label {
                font-size: 1rem;
                color: #31333F;
                white-space: nowrap;
            }
            </style>
        """, unsafe_allow_html=True)

        # 상단 메트릭 표시
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
                <div class="metric-container">
                    <div class="amount-label">기말잔액</div>
                    <div class="amount-box">
                        <div class="amount-value">{final_balance:,.0f}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
                <div class="metric-container">
                    <div class="amount-label">입금</div>
                    <div class="amount-box">
                        <div class="amount-value">{total_deposit:,.0f}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
                <div class="metric-container">
                    <div class="amount-label">출금</div>
                    <div class="amount-box">
                        <div class="amount-value">{total_withdrawal:,.0f}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

        # 2. 계좌별 통합 현황 제목 (여백 추가)
        st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)  # 상단 여백 추가
        st.header("2. 계좌별 통합 현황")
        
        # 자금일보 시트 전체 읽기
        df_raw = pd.read_excel(uploaded_file, sheet_name="자금일보", header=None)
        
        # 1. "계좌별 통합 현황" 위치 찾기
        summary_row = None
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.contains("계좌별 통합 현황").any():
                summary_row = idx
                break
        
        if summary_row is not None:
            # 2. 헤더 행 찾기
            header_row = summary_row + 1
            
            # 3. 실제 데이터 시작 행
            data_start_row = header_row + 1
            
            # 4. 합계 행 찾기
            end_row = None
            for idx, row in df_raw.iloc[data_start_row:].iterrows():
                if row.astype(str).str.contains("합계").any():
                    end_row = idx
                    break
            
            # 5. 계좌 정보 읽기
            account_info = pd.read_excel(
                uploaded_file,
                sheet_name="자금일보",
                header=header_row,
                nrows=end_row-data_start_row
            )
            
            # 6. 각 계좌별 잔액 계산
            results = []
            for _, row in account_info.iterrows():
                account_no = row['계좌번호']
                
                # 기초잔액 계산
                initial_balance = float(df_daily[
                    (df_daily['지출일'] < start_datetime) &
                    (df_daily['계좌번호'] == account_no)
                ]['집행 금액'].sum())
                
                # 입금 계산 ("계좌 대체" 제외)
                deposits = df_daily[
                    (df_daily['지출일'] >= start_datetime) &
                    (df_daily['지출일'] <= end_datetime) &
                    (df_daily['계좌번호'] == account_no) &
                    (df_daily['입금'].notna()) &
                    (df_daily['현금흐름 대분류'] != "계좌 대체")  # 계좌 대체 제외
                ]['입금'].sum()
                
                # 출금 계산 ("계좌 대체" 제외)
                withdrawals = df_daily[
                    (df_daily['지출일'] >= start_datetime) &
                    (df_daily['지출일'] <= end_datetime) &
                    (df_daily['계좌번호'] == account_no) &
                    (df_daily['출금'].notna()) &
                    (df_daily['현금흐름 대분류'] != "계좌 대체")  # 계좌 대체 제외
                ]['출금'].sum()
                
                # 기말잔액 계산
                final_balance = initial_balance + deposits - withdrawals
                
                results.append({
                    '구분': row['구분'],
                    '금융사': row['금융사'],
                    '계좌번호': account_no,
                    '기초잔액': initial_balance,
                    '입금': deposits,
                    '출금': -withdrawals,  # 출금은 음수로 표시
                    '기말잔액': final_balance
                })
            
            # 7. 결과를 데이터프레임으로 변환
            summary_df = pd.DataFrame(results)
            
            # 숫자 컬럼들을 float 타입으로 변환
            numeric_columns = ['기초잔액', '입금', '출금', '기말잔액']
            for col in numeric_columns:
                summary_df[col] = pd.to_numeric(summary_df[col], errors='coerce')
            
            # 8. 합계 행 추가
            totals = summary_df.agg({
                '기초잔액': 'sum',
                '입금': 'sum',
                '출금': 'sum',
                '기말잔액': 'sum'
            }).to_frame().T
            totals['구분'] = '합계'
            totals['금융사'] = ''
            totals['계좌번호'] = ''
            
            final_df = pd.concat([summary_df, totals], ignore_index=True)
            
            # 최종 데이터프레임의 숫자 컬럼들도 float로 변환
            for col in numeric_columns:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

                # 데이터프레임 표시하기 전에 숫자 컬럼들을 먼저 변환
                numeric_columns = ['기초잔액', '입금', '출금', '기말잔액']
                
                # 숫자 포맷팅 함수
                def format_number(x):
                    if pd.isna(x):
                        return ''
                    return '{:,.0f}'.format(float(x))

                # 숫자 컬럼들을 float로 변환 후 포맷팅
                final_df_display = final_df.copy()
                for col in numeric_columns:
                    final_df_display[col] = pd.to_numeric(final_df_display[col], errors='coerce')
                    final_df_display[col] = final_df_display[col].apply(format_number)

                # 테이블 표시
            st.dataframe(
                    final_df_display,  # 포맷팅된 데이터프레임 사용
                hide_index=True,
                column_config={
                    "구분": st.column_config.TextColumn("구분", width="medium"),
                    "금융사": st.column_config.TextColumn("금융사", width="medium"),
                    "계좌번호": st.column_config.TextColumn("계좌번호", width="medium"),
                        "기초잔액": st.column_config.TextColumn("기초잔액", width="medium"),
                        "입금": st.column_config.TextColumn("입금", width="medium"),
                        "출금": st.column_config.TextColumn("출금", width="medium"),
                        "기말잔액": st.column_config.TextColumn("기말잔액", width="medium")
                },
                use_container_width=True
            )

        # 3. 입출금 상세 내역 (버튼 없이 바로 표시)
        st.header("3. 입출금 상세 내역")
        
        
            # 전체 데이터 준비 (그래프용)
        df_all_transactions = df_daily[
            (df_daily["지출일"] > start_datetime) &
            (df_daily["지출일"] <= end_datetime) &
            (df_daily["현금흐름 대분류"] != "계좌 대체")
            ].copy()

            # 입금 상세내역
        st.subheader("입금 상세내역")
        details = df_all_transactions[df_all_transactions["입금"].notna()]
        
        if not details.empty:
            display_columns = ["금융사", "계좌번호", "지출일", "집행 금액", "적요", 
                            "현금흐름 대분류", "현금흐름 중분류"]
            details_display = details[display_columns].copy()
                
                # 집행 금액 포맷팅
            details_display['집행 금액'] = details_display['집행 금액'].apply(
                    lambda x: '{:,.0f}'.format(float(x)) if pd.notna(x) else ''
                )
            
            # 합계 행 추가
            total_amount = details['집행 금액'].sum()
            totals = pd.DataFrame([{
                '금융사': '합계',
                '계좌번호': '',
                '지출일': None,
                    '집행 금액': '{:,.0f}'.format(float(total_amount)),
                '적요': '',
                '현금흐름 대분류': '',
                '현금흐름 중분류': ''
            }])
            details_display = pd.concat([details_display, totals])
            
            st.dataframe(
                details_display,
                hide_index=True,
                use_container_width=True
            )

                # 입금 현금흐름 중분류별 분석
            st.subheader("입금 항목별 분석")
                
                # 데이터 확인 및 처리
            details['집행 금액_숫자'] = pd.to_numeric(details['집행 금액'], errors='coerce')
                
                # 중분류별 금액 집계
            inflow_by_category = details.groupby('현금흐름 중분류')['집행 금액_숫자'].sum()
                
                # 데이터프레임으로 변환하고 컬럼명 변경
            inflow_df = inflow_by_category.reset_index()
            inflow_df.columns = ['중분류', '금액']
                
                # 금액 포맷팅
            inflow_df['금액'] = inflow_df['금액'].apply(lambda x: '{:,.0f}'.format(x))
                
                # 입금용 색상 팔레트 (붉은색, 주황색, 노란색 계열)
            inflow_colors = [
                    '#FF6B6B',  # 붉은색
                    '#FFA07A',  # 연한 주황
                    '#FFB74D',  # 진한 주황
                    '#FFD700',  # 골든로드
                    '#FFF176',  # 연한 노랑
                    '#FF7043',  # 깊은 주황
                    '#FF9800',  # 표준 주황
                    '#FFEB3B'   # 밝은 노랑
                ]

                # 출금용 색상 팔레트 (푸른색, 초록색 계열)
            outflow_colors = [
                    '#4CAF50',  # 초록
                    '#2196F3',  # 파랑
                    '#00BCD4',  # 청록
                    '#009688',  # 틸
                    '#3F51B5',  # 남색
                    '#8BC34A',  # 연한 초록
                    '#03A9F4',  # 연한 파랑
                    '#00796B'   # 깊은 청록
                ]

                # 테이블과 도넛 차트를 나란히 배치
            col1, col2 = st.columns([1, 2])  # 1:2 비율로 분할
                
            with col1:
                    # 중분류별 금액 테이블
                    st.write("중분류별 금액:")
                    st.dataframe(
                        inflow_df,
                        hide_index=True,
                        column_config={
                            "중분류": st.column_config.TextColumn(
                                "중분류",
                                width="small"
                            ),
                            "금액": st.column_config.TextColumn(
                                "금액",
                                width="small"
                            )
                        },
                        use_container_width=True
                    )
                
            with col2:
                    # 파이 차트 (입금용)
                    fig_pie = go.Figure(data=[go.Pie(
                        labels=inflow_by_category.index,
                        values=inflow_by_category.values,
                        hole=.3,
                        hovertemplate="<b>%{label}</b><br>" +
                                    "금액: ￦%{value:,.0f}<br>" +
                                    "비중: %{percent}<extra></extra>",
                        textposition="outside",  # 라벨을 외부에 표시
                        textinfo="label+percent",  # 라벨과 퍼센트 모두 표시
                        showlegend=True,  # 범례 표시
                        marker_colors=inflow_colors  # 입금용 색상 적용
                    )])
                    fig_pie.update_layout(
                        title="입금 항목별 비중",
                        height=400,
                        margin=dict(t=30, b=30, l=50, r=100),  # 우측 여백 증가
                        annotations=[dict(
                            text='입금<br>비중',
                            x=0.5,
                            y=0.5,
                            font_size=15,
                            showarrow=False
                        )],
                        legend=dict(
                            yanchor="top",
                            y=1.0,
                            xanchor="left",
                            x=1.02
                        )
                    )
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                # 입금 막대 그래프
            fig_bar = go.Figure(data=[go.Bar(
                    x=inflow_by_category.index,
                    y=inflow_by_category.values,
                    text=[f'￦{x:,.0f}' for x in inflow_by_category.values],
                    textposition='auto',
                    marker_color=inflow_colors  # 입금용 색상 적용
                )])
            fig_bar.update_layout(
                    title="입금 항목별 금액",
                    yaxis_title="금액(원)",
                    height=400,
                    yaxis=dict(tickformat=",")
                )
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.write("해당 기간에 입금 내역이 없습니다.")

        # 출금 상세내역
        st.subheader("출금 상세내역")
        details = df_all_transactions[df_all_transactions["출금"].notna()]
        
        if not details.empty:
            display_columns = ["금융사", "계좌번호", "지출일", "집행 금액", "적요", 
                            "현금흐름 대분류", "현금흐름 중분류"]
            details_display = details[display_columns].copy()
                
                # 집행 금액 포맷팅
            details_display['집행 금액'] = details_display['집행 금액'].apply(
                    lambda x: '{:,.0f}'.format(float(x)) if pd.notna(x) else ''
                )
            
            # 합계 행 추가
            total_amount = details['집행 금액'].sum()
            totals = pd.DataFrame([{
                '금융사': '합계',
                '계좌번호': '',
                '지출일': None,
                    '집행 금액': '{:,.0f}'.format(float(total_amount)),
                '적요': '',
                '현금흐름 대분류': '',
                '현금흐름 중분류': ''
            }])
            details_display = pd.concat([details_display, totals])
            
            st.dataframe(
                details_display,
                hide_index=True,
                use_container_width=True
            )
                
                # 출금 현금흐름 중분류별 분석 그래프
            st.subheader("출금 항목별 분석")
                
                # 데이터 확인 및 처리
            details['집행 금액_숫자'] = pd.to_numeric(details['집행 금액'], errors='coerce').abs()
                
                # 중분류별 금액 집계
            outflow_by_category = details.groupby('현금흐름 중분류')['집행 금액_숫자'].sum()
                
                # 데이터프레임으로 변환하고 컬럼명 변경
            outflow_df = outflow_by_category.reset_index()
            outflow_df.columns = ['중분류', '금액']
                
                # 금액 포맷팅
            outflow_df['금액'] = outflow_df['금액'].apply(lambda x: '{:,.0f}'.format(x))
                
                # 테이블과 도넛 차트를 나란히 배치
            col1, col2 = st.columns([1, 2])  # 1:2 비율로 분할
                
            with col1:
                    # 데이터 확인용 출력 (디버깅)
                    st.write("중분류별 금액:")
                    st.dataframe(
                        outflow_df,
                        hide_index=True,
                        column_config={
                            "중분류": st.column_config.TextColumn(
                                "중분류",
                                width="small"
                            ),
                            "금액": st.column_config.TextColumn(
                                "금액",
                                width="small"
                            )
                        },
                        use_container_width=True
                    )
                
            with col2:
                    # 파이 차트 (출금용)
                    fig_pie = go.Figure(data=[go.Pie(
                        labels=outflow_by_category.index,
                        values=outflow_by_category.values,
                        hole=.3,
                        hovertemplate="<b>%{label}</b><br>" +
                                    "금액: ￦%{value:,.0f}<br>" +
                                    "비중: %{percent}<extra></extra>",
                        textposition="outside",  # 라벨을 외부에 표시
                        textinfo="label+percent",  # 라벨과 퍼센트 모두 표시
                        showlegend=True,  # 범례 표시
                        marker_colors=outflow_colors  # 출금용 색상 적용
                    )])
                    fig_pie.update_layout(
                        title="출금 항목별 비중",
                        height=400,
                        margin=dict(t=30, b=30, l=50, r=100),  # 우측 여백 증가
                        annotations=[dict(
                            text='출금<br>비중',
                            x=0.5,
                            y=0.5,
                            font_size=15,
                            showarrow=False
                        )],
                        legend=dict(
                            yanchor="top",
                            y=1.0,
                            xanchor="left",
                            x=1.02
                        )
                    )
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                # 출금 막대 그래프
            fig_bar = go.Figure(data=[go.Bar(
                    x=outflow_by_category.index,
                    y=-outflow_by_category.values,
                    text=[f'￦{x:,.0f}' for x in outflow_by_category.values],
                    textposition='auto',
                    marker_color=outflow_colors  # 출금용 색상 적용
                )])
            fig_bar.update_layout(
                    title="출금 항목별 금액",
                    yaxis_title="금액(원)",
                    height=400,
                    yaxis=dict(tickformat=",")
                )
            st.plotly_chart(fig_bar, use_container_width=True)

            # 일별 입출금 추이 그래프 추가
            if not df_all_transactions.empty:
                st.subheader("일별 입출금 추이")
                
                daily_summary = df_all_transactions.groupby("지출일").agg({
                    "입금": "sum",
                    "출금": "sum"
                }).reset_index()
                
                fig = go.Figure()
                
                # 입금 막대 그래프
                fig.add_trace(go.Bar(
                    name='입금',
                    x=daily_summary['지출일'],
                    y=daily_summary['입금'].fillna(0),
                    marker_color='rgba(244, 67, 54, 0.7)',  # 부드러운 빨간색
                    hovertemplate='<b>입금</b>: %{y:,.0f}원<extra></extra>'
                ))
                
                # 출금 막대 그래프
                fig.add_trace(go.Bar(
                    name='출금',
                    x=daily_summary['지출일'],
                    y=-daily_summary['출금'].fillna(0),
                    marker_color='rgba(33, 150, 243, 0.7)',  # 부드러운 파란색
                    hovertemplate='<b>출금</b>: %{y:,.0f}원<extra></extra>'
                ))
                
                # 순변동 선 그래프
                net_change = daily_summary['입금'].fillna(0) - daily_summary['출금'].fillna(0)
                fig.add_trace(go.Scatter(
                    name='순변동',
                    x=daily_summary['지출일'],
                    y=net_change,
                    line=dict(color='rgba(156, 39, 176, 0.9)', width=3),  # 보라색
                    mode='lines+markers',  # 선과 점을 함께 표시
                    marker=dict(size=8),
                    hovertemplate='<b>순변동</b>: %{y:,.0f}원<extra></extra>'
                ))
                
                # 그래프 레이아웃 설정
                fig.update_layout(
                    barmode='relative',
                    title={
                        'text': '일별 입출금 현황',
                        'y':0.95,
                        'x':0.5,
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': dict(size=20)
                    },
                    xaxis_title='거래일자',
                    yaxis_title='금액(원)',
                    height=500,  # 그래프 높이 증가
                    hovermode='x unified',
                    plot_bgcolor='rgba(255,255,255,0.9)',  # 배경색 설정
                    paper_bgcolor='rgba(255,255,255,0.9)',
                    yaxis=dict(
                        tickformat=',',
                        gridcolor='rgba(0,0,0,0.1)',  # 그리드 색상
                        zerolinecolor='rgba(0,0,0,0.2)',  # 0선 색상
                    ),
                    xaxis=dict(
                        gridcolor='rgba(0,0,0,0.1)',
                        showgrid=True,
                    ),
                    showlegend=True,
                    legend=dict(
                        orientation='h',
                        yanchor='bottom',
                        y=1.02,
                        xanchor='right',
                        x=1,
                        bgcolor='rgba(255,255,255,0.8)',
                        bordercolor='rgba(0,0,0,0.2)',
                        borderwidth=1
                    ),
                    margin=dict(l=50, r=50, t=80, b=50)  # 여백 조정
                )
                
                # 그래프 테마 업데이트
                fig.update_xaxes(showline=True, linewidth=1, linecolor='rgba(0,0,0,0.2)')
                fig.update_yaxes(showline=True, linewidth=1, linecolor='rgba(0,0,0,0.2)')
                
                st.plotly_chart(fig, use_container_width=True)

    # 두 번째 탭: 현금흐름표
    with tab2:
        st.header("현금흐름표")
        
        # 시작월과 종료월 선택
        container = st.container()
        with container:
            col1, col2 = st.columns([2, 8])  # 좌측에 더 작은 공간 할당
            
            with col1:  # 좌측 컬럼에 선택 박스 배치
                subcol1, subcol2 = st.columns(2)
                with subcol1:
                    start_year = st.selectbox("시작 연도", range(2020, 2031), index=4)
                    start_month = st.selectbox("시작 월", range(1, 13), index=1)
                
                with subcol2:
                    end_year = st.selectbox("종료 연도", range(2020, 2031), index=4)
                    end_month = st.selectbox("종료 월", range(1, 13), index=4)

        # 합계 0인 행 숨기기 체크박스 추가
        hide_zero_rows = st.checkbox("합계 0인 행 숨기기")

        if uploaded_file is not None:
            # 기본 헤더 정의
            base_headers = ['Level', '현금 흐름 구분', '유입/유출', '구분1', '구분2', 'CODE']
            
            # 전체 데이터 읽기
            df_full = pd.read_excel(
                uploaded_file,
                sheet_name="월별_CashFlow",
                skiprows=2
            )
            
            # Level 1인 마지막 행 찾기 (실제 데이터의 끝 지점)
            last_row_mask = (df_full['Level'] == 1)
            matching_rows = df_full.index[last_row_mask].tolist()
            
            if matching_rows:
                # 마지막 Level 1 행의 인덱스가 아닌, 전체 Level 1 행을 포함하도록 수정
                df_base = df_full.iloc[:matching_rows[-1] + 1, 3:9]  # D열부터 I열까지 (인덱스는 0부터 시작하므로 3:9)
            else:
                st.error("데이터에서 Level 1인 행을 찾을 수 없습니다.")
                df_base = df_full.iloc[:, 3:9]
            
            # 날짜 데이터 처리
            date_cols = df_full.columns[9:]  # J열부터의 열 이름들
            date_mapping = {}
            
            for col in date_cols:
                try:
                    if pd.notna(col):
                        date = pd.to_datetime(col)
                        formatted_date = f"{date.year}년 {date.month:02d}월"
                        date_mapping[col] = formatted_date
                except:
                    continue
            
            # 선택된 기간의 월별 열 생성
            selected_months = []
            current_date = pd.Timestamp(f"{start_year}-{start_month}-01")
            end_date = pd.Timestamp(f"{end_year}-{end_month}-01")
            
            while current_date <= end_date:
                selected_months.append(f"{current_date.year}년 {current_date.month:02d}월")
                current_date = current_date + pd.DateOffset(months=1)
            
            # 최종 데이터프레임 생성
            result_columns = base_headers + selected_months + ['합계']  # '증감'을 '합계'로 변경
            df_result = pd.DataFrame(columns=result_columns)
            df_result[base_headers] = df_base
            
            # 선택된 기간의 데이터 매핑
            monthly_data = df_full.iloc[:matching_rows[-1] + 1, 9:]  # J열부터의 데이터
            for orig_col, formatted_date in date_mapping.items():
                if formatted_date in selected_months:
                    df_result[formatted_date] = monthly_data[orig_col]
            
            # None 값을 빈 문자열로 변경 (특정 열만)
            text_columns = ['현금 흐름 구분', '유입/유출', '구분1', '구분2']
            df_result[text_columns] = df_result[text_columns].fillna('')
            
            # 6. 합계 계산
            if len(selected_months) > 0:
                # 합계 0인 행 숨기기 기능
                if hide_zero_rows:
                    sum_columns = selected_months
                    row_sums = df_result[sum_columns].astype(float).sum(axis=1)
                    mask = (df_result['Level'].isin([1, 2])) | (row_sums != 0)
                    df_result = df_result[mask]
                
                # 선택된 모든 월의 합계 계산
                df_result['합계'] = df_result[selected_months].astype(float).sum(axis=1)
                df_result['합계'] = df_result['합계'].apply(lambda x: f'{x:,.0f}' if pd.notnull(x) else '')
                
                # 기초현금과 기말현금 행의 합계는 빈 문자열로 설정
                df_result.loc[df_result['현금 흐름 구분'] == '기초현금', '합계'] = ''
                df_result.loc[df_result['현금 흐름 구분'] == '기말현금', '합계'] = ''

                # 표 스타일링
                def color_rows(row):
                    level = row['Level']
                    if level == 1:
                        return ['background-color: #E5E5E5'] * len(row)
                    elif level == 2:
                        return ['background-color: rgb(235, 244, 245)'] * len(row)
                    else:
                        return [''] * len(row)

                styled_df = df_result.style\
                    .format(lambda x: f'{x:,.0f}' if isinstance(x, (int, float)) else x)\
                    .apply(color_rows, axis=1)

            # 결과 표시
            st.write("현금흐름표:")
            st.dataframe(
                styled_df,
                use_container_width=True,
                height=len(df_result) * 35 + 38
            )

            # 그래프를 위한 데이터 준비
            if len(selected_months) > 0:
                try:
                    # 기초현금과 기말현금 데이터 추출
                    initial_cash_row = df_result[df_result['현금 흐름 구분'] == '기초현금']
                    final_cash_row = df_result[df_result['현금 흐름 구분'] == '기말현금']
                    
                    if not initial_cash_row.empty and not final_cash_row.empty:
                        initial_cash = initial_cash_row[selected_months].values[0].astype(float)
                        final_cash = final_cash_row[selected_months].values[0].astype(float)

                        # 1. 월별 현금 잔액 추이 그래프
                        fig1 = go.Figure()
                        fig1.add_trace(go.Scatter(
                            x=selected_months, 
                            y=initial_cash,
                            name='기초현금', 
                            line=dict(color='green', width=2)
                        ))
                        fig1.add_trace(go.Scatter(
                            x=selected_months, 
                            y=final_cash,
                            name='기말현금', 
                            line=dict(color='red', dash='dash')
                        ))
                        fig1.update_layout(
                            title='월별 현금 잔액 추이',
                            xaxis_title='연도월',
                            yaxis_title='금액',
                            height=400,
                            yaxis=dict(tickformat=",")
                        )
                        st.plotly_chart(fig1, use_container_width=True)

                        # 2. 현금 유입/유출 비교 그래프
                        cash_flows = {
                            '영업활동': [],
                            '투자활동': [],
                            '재무활동': []
                        }
                        
                        # Level 1인 행에서 영업, 투자, 재무 데이터 추출
                        for keyword, activity in [('영업', '영업활동'), ('투자', '투자활동'), ('재무', '재무활동')]:
                            mask = (df_result['Level'] == 1) & (df_result['현금 흐름 구분'].str.contains(keyword, na=False, case=False))
                            matching_rows = df_result[mask]
                            
                            if not matching_rows.empty:
                                # 선택된 월의 데이터를 리스트로 변환
                                cash_flows[activity] = matching_rows[selected_months].values[0].astype(float)

                        # 데이터가 있는 경우에만 그래프 생성
                        if any(len(flows) > 0 for flows in cash_flows.values()):
                            fig2 = go.Figure()
                            colors = {
                                '영업활동': 'rgba(244, 67, 54, 0.7)',    # 부드러운 빨간색
                                '투자활동': 'rgba(33, 150, 243, 0.7)',   # 부드러운 파란색
                                '재무활동': 'rgba(156, 39, 176, 0.7)'    # 보라색
                            }
                            
                            for activity, flows in cash_flows.items():
                                if len(flows) > 0:  # 데이터가 있는 경우만 그래프에 추가
                                    fig2.add_trace(go.Bar(
                                        name=activity,
                                        x=selected_months,
                                        y=flows,
                                        marker_color=colors[activity]
                                    ))
                            
                            fig2.update_layout(
                                title='현금 유입/유출 비교',
                                xaxis_title='연도월',
                                yaxis_title='금액',
                                barmode='group',
                                height=400,
                                showlegend=True,
                                legend_title='활동 구분',
                                yaxis=dict(tickformat=",")
                            )
                            st.plotly_chart(fig2, use_container_width=True)

                    # 변동비/고정비 상세 비중 분석
                    try:
                        # 변동비와 고정비 행 찾기
                        variable_cost_rows = df_result[df_result['구분1'] == '변동비']
                        fixed_cost_rows = df_result[df_result['구분1'] == '고정비']

                        if not variable_cost_rows.empty and not fixed_cost_rows.empty:
                            variable_cost_idx = variable_cost_rows.index[0]
                            fixed_cost_idx = fixed_cost_rows.index[0]

                            # 변동비 하위레벨 추출
                            variable_subcosts = df_result[
                                (df_result.index > variable_cost_idx) & 
                                (df_result.index < fixed_cost_idx) &
                                (df_result['구분2'].notna())
                            ]
                            
                            # 금액이 0이 아닌 행만 선택
                            variable_subcosts = variable_subcosts[
                                variable_subcosts[selected_months].astype(float).abs().sum(axis=1) > 0
                            ]

                            # 고정비 하위레벨 추출
                            fixed_subcosts = df_result[
                                (df_result.index > fixed_cost_idx) & 
                                (df_result['Level'] == 4)
                            ]
                            
                            # Level이 3이나 1이 나오는 행 이후는 제외
                            if not fixed_subcosts.empty:
                                end_idx = fixed_subcosts.index[-1]
                                for idx in fixed_subcosts.index:
                                    if df_result.loc[idx:, 'Level'].isin([1, 3]).any():
                                        end_idx = df_result.loc[idx:, 'Level'].isin([1, 3]).idxmax()
                                        break
                                fixed_subcosts = fixed_subcosts.loc[:end_idx-1]
                            
                            # 금액이 0이 아닌 행만 선택
                            fixed_subcosts = fixed_subcosts[
                                fixed_subcosts[selected_months].astype(float).abs().sum(axis=1) > 0
                            ]

                            # 그래프 생성
                            fig_cost = go.Figure()

                            # 색상 정의
                            orange_colors = [
                                'rgba(255, 87, 34, 0.7)',   # 진한 주황
                                'rgba(255, 152, 0, 0.7)',   # 주황
                                'rgba(255, 193, 7, 0.7)',   # 황색
                                'rgba(255, 235, 59, 0.7)',  # 연한 황색
                                'rgba(251, 140, 0, 0.7)'    # 다크 주황
                            ]

                            blue_colors = [
                                'rgba(33, 150, 243, 0.7)',   # 진한 파랑
                                'rgba(3, 169, 244, 0.7)',    # 파랑
                                'rgba(0, 188, 212, 0.7)',    # 연한 파랑
                                'rgba(178, 235, 242, 0.7)',  # 매우 연한 파랑
                                'rgba(21, 101, 192, 0.7)'    # 다크 파랑
                            ]

                            # 변동비 하위 항목 추가
                            for idx, (_, row) in enumerate(variable_subcosts.iterrows()):
                                fig_cost.add_trace(go.Bar(
                                    name=f'변동비-{row["구분2"]}',
                                    x=selected_months,
                                    y=row[selected_months].astype(float).abs(),
                                    marker_color=orange_colors[idx % len(orange_colors)],
                                    text=[f'￦{abs(float(v)):,.0f}' for v in row[selected_months]],
                                    textposition='auto',
                                    legendgroup='변동비',
                                    legendgrouptitle_text='변동비'
                                ))

                            # 고정비 하위 항목 추가
                            for idx, (_, row) in enumerate(fixed_subcosts.iterrows()):
                                fig_cost.add_trace(go.Bar(
                                    name=f'고정비-{row["구분2"]}',
                                    x=selected_months,
                                    y=row[selected_months].astype(float).abs(),
                                    marker_color=blue_colors[idx % len(blue_colors)],
                                    text=[f'￦{abs(float(v)):,.0f}' for v in row[selected_months]],
                                    textposition='auto',
                                    legendgroup='고정비',
                                    legendgrouptitle_text='고정비'
                                ))

                            # 총액 선 그래프 추가
                            total_variable = variable_subcosts[selected_months].astype(float).abs().sum()
                            total_fixed = fixed_subcosts[selected_months].astype(float).abs().sum()

                            fig_cost.add_trace(go.Scatter(
                                name='변동비 총액',
                                x=selected_months,
                                y=total_variable,
                                line=dict(color='rgba(255, 87, 34, 1)', width=2),
                                legendgroup='변동비'
                            ))

                            fig_cost.add_trace(go.Scatter(
                                name='고정비 총액',
                                x=selected_months,
                                y=total_fixed,
                                line=dict(color='rgba(33, 150, 243, 1)', width=2),
                                legendgroup='고정비'
                            ))

                            # 레이아웃 설정
                            fig_cost.update_layout(
                                title='월별 변동비/고정비 상세 내역',
                                barmode='stack',
                                height=500,
                                yaxis=dict(
                                    title='금액(원)',
                                    tickformat=','
                                ),
                                showlegend=True,
                                legend=dict(
                                    groupclick="toggleitem"
                                )
                            )

                            st.plotly_chart(fig_cost, use_container_width=True)

                        # 표 생성을 위한 데이터 준비
                        summary_data = {
                            '구분': ['변동비 합계', '고정비 합계', '변동비 비중(%)', '고정비 비중(%)', '총계'],
                        }
                        
                        # 각 월별 금액과 비중 계산
                        for month in selected_months:
                            variable_sum = variable_subcosts[month].astype(float).abs().sum()
                            fixed_sum = fixed_subcosts[month].astype(float).abs().sum()
                            total_sum = variable_sum + fixed_sum
                            
                            # 비중 계산 (총계가 0인 경우 예외 처리)
                            if total_sum > 0:
                                variable_ratio = (variable_sum / total_sum) * 100
                                fixed_ratio = (fixed_sum / total_sum) * 100
                            else:
                                variable_ratio = fixed_ratio = 0
                            
                            summary_data[month] = [
                                f'{variable_sum:,.0f}',
                                f'{fixed_sum:,.0f}',
                                f'{variable_ratio:.1f}%',
                                f'{fixed_ratio:.1f}%',
                                f'{total_sum:,.0f}'
                            ]

                        # DataFrame 생성
                        summary_df = pd.DataFrame(summary_data)
                        
                        # 표 스타일링
                        def style_summary_table(df):
                            return df.style \
                                .set_properties(**{
                                    'text-align': 'right',
                                    'font-size': '14px',
                                    'padding': '5px 15px'
                                }) \
                                .set_properties(subset=['구분'], **{
                                    'text-align': 'left',
                                    'font-weight': 'bold'
                                }) \
                                .set_table_styles([
                                    {'selector': 'th',
                                     'props': [('text-align', 'center'),
                                             ('font-weight', 'bold'),
                                             ('background-color', '#f0f2f6'),
                                             ('padding', '5px 15px')]},
                                    {'selector': 'td',
                                     'props': [('border', '1px solid #ddd')]},
                                    {'selector': 'th',
                                     'props': [('border', '1px solid #ddd')]}
                                ])

                        # 표 출력
                        st.markdown("#### 월별 변동비/고정비 요약")
                        st.dataframe(style_summary_table(summary_df), hide_index=True)

                        # 기말 현금잔액 예상액 추이 그래프 생성
                        try:
                            # 숫자 열만 선택 (현금 흐름 구분 등의 문자열 열 제외)
                            numeric_columns = [col for col in df_result.columns if '년' in str(col) and '월' in str(col)]
                            # 마지막 열을 종료월로 설정
                            end_month = numeric_columns[-1]
                            
                            # 최근 3개월 선택
                            recent_months = sorted(numeric_columns)[-3:]
                            
                            # 고정비 합계 행 찾기
                            fixed_cost_row = summary_df[summary_df['구분'] == '고정비 합계']
                            
                            # 최근 3개월 고정비 평균 계산
                            recent_fixed_costs = []
                            for month in recent_months:
                                cost_str = fixed_cost_row[month].iloc[0]
                                # 문자열에서 숫자로 변환 ('￦' 기호와 쉼표 제거)
                                cost = float(str(cost_str).replace('￦', '').replace(',', ''))
                                recent_fixed_costs.append(cost)
                            
                            avg_fixed_cost = sum(recent_fixed_costs) / len(recent_fixed_costs)
                            st.write("5. 평균 고정비 (3개월 평균 고정비):", f"{avg_fixed_cost:,.0f}")
                            
                            # 기말현금 데이터 확인
                            ending_cash_row = df_result[df_result['현금 흐름 구분'] == '기말현금']
                            
                            if not ending_cash_row.empty:
                                # 초기 현금 확인
                                initial_cash_str = ending_cash_row[end_month].iloc[0]
                                
                                # 문자열 처리
                                if isinstance(initial_cash_str, str):
                                    initial_cash = float(initial_cash_str.replace(',', ''))
                                else:
                                    initial_cash = float(initial_cash_str)
                                st.write("6. 초기 현금:", f"{initial_cash:,.0f}")
                                
                                # 미래 현금 계산
                                future_cash = []
                                dates = []
                                current_cash = initial_cash
                                
                                # 종료 연월 파싱
                                end_year = int(end_month.split('년')[0])
                                end_month_num = int(end_month.split('년')[1].split('월')[0])
                                
                                # 현금이 음수가 될 때까지 계산 (24개월 제한 삭제)
                                month_count = 0
                                while current_cash > 0:  # 24개월 제한 조건 제거
                                    future_cash.append(current_cash)
                                    
                                    new_month = end_month_num + month_count
                                    year = end_year + (new_month - 1) // 12
                                    month = ((new_month - 1) % 12) + 1
                                    dates.append(f'{year}년 {month:02d}월')
                                    
                                    current_cash -= avg_fixed_cost
                                    month_count += 1
                                
                                # 마지막 음수 값 포함
                                future_cash.append(current_cash)
                                new_month = end_month_num + month_count
                                year = end_year + (new_month - 1) // 12
                                month = ((new_month - 1) % 12) + 1
                                dates.append(f'{year}년 {month:02d}월')
                                
                                # 그래프 생성
                                fig_forecast = go.Figure()
                                
                                fig_forecast.add_trace(go.Scatter(
                                    x=dates,
                                    y=future_cash,
                                    mode='lines+markers+text',
                                    name='예상 기말현금',
                                    line=dict(color='rgb(33, 150, 243)', width=2),
                                    text=[f'￦{val:,.0f}' for val in future_cash],
                                    textposition='top center'
                                ))
                                
                                fig_forecast.update_layout(
                                    title='기말 현금잔액 예상액 추이',
                                    xaxis_title='연월',
                                    yaxis_title='금액(원)',
                                    yaxis=dict(tickformat=','),
                                    showlegend=True,
                                    height=500
                                )
                                
                                fig_forecast.add_hline(y=0, line_dash="dash", line_color="red")
                                
                                st.plotly_chart(fig_forecast, use_container_width=True)
                                
                                # 현금 소진 예상 시점 안내 (항상 표시)
                                for i, cash in enumerate(future_cash):
                                    if cash <= 0:
                                        st.warning(f'현재 고정비 지출 수준 유지 시 {dates[i]}에 현금이 소진될 것으로 예상됩니다.')
                                        break

                        except Exception as e:
                            st.error(f"오류 발생 위치 확인: {str(e)}")
                            import traceback
                            st.error(f"상세 오류: {traceback.format_exc()}")

                    except Exception as e:
                        st.error("변동비/고정비 상세 분석 그래프 생성 중 오류가 발생했습니다.")
                except Exception as e:
                    st.error(f"그래프 생성 중 오류 발생: {str(e)}")
                    st.write("오류 상세:", e)

                    