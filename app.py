import streamlit as st
import pandas as pd
import numpy as np
import io
from model.intern_assign import WORKFORCE_ASSIGN # 최적화 코드 
from model.make_excel import create_excel_file 

# -----------------------------------------------------------------------------
# 1. 초기 설정 (1920x1080 고정)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Workforce Planner",
    page_icon="📆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# 2. 상태 관리
# -----------------------------------------------------------------------------
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

def reset_uploader():
    st.session_state['uploader_key'] += 1

# -----------------------------------------------------------------------------
# 3. CSS 스타일
# -----------------------------------------------------------------------------
def set_dashboard_style():
    st.markdown("""
        <style>
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
        html, body, [class*="css"] {
            font-family: Pretendard, -apple-system, sans-serif !important;
            font-size: 18px !important;
        }

        header[data-testid="stHeader"] { display: none !important; }
        
        html, body, .stApp {
            overflow: hidden !important;
            /* background-color: #F3F4F6; */ /* 기존 배경 유지 */
        }

        .block-container {
            padding-top: 1.5rem !important; 
            padding-bottom: 2rem !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
            max-width: 1920px !important; 
            min-width: 1920px !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
            overflow: hidden !important;
            margin: 0 auto;
        }
        
        /* [강제 고정 모드] */
        /* 1. 줄바꿈 원천 차단 */
        div[data-testid="stHorizontalBlock"] {
            width: 100% !important;
            flex-wrap: nowrap !important;
            gap: 1rem !important;
        }
        
        /* 2. 컬럼 너비 강제 유지 */
        div[data-testid="column"] {
            flex: 1 1 auto !important;
            min-width: auto !important;
        }
        
        /* 3. 전체 화면 강제 확장 (스크롤 생성 유도 -> 스크롤 숨김) */
        html, body, .stApp {
            min-width: 1920px !important;
            overflow: hidden !important;
        }


        /* [카드 레이아웃] */
        .card-title {
            font-size: 1.2rem; font-weight: 700; color: #111827;
            display: flex; align-items: center; gap: 8px;
            white-space: nowrap;
        }

        /* [업로더 스타일] */
        [data-testid="stFileUploader"] {
            background-color: #EFF6FF;
            border: 2px dashed #3B82F6;
            border-radius: 8px;
            padding: 0px; 
            text-align: center;
            min-height: 80px; 
            display: flex; align-items: center; justify-content: center;
        }
        [data-testid="stFileUploader"] section { 
            padding: 10px !important; min-height: 0px !important; width: 100% !important;
        }
        [data-testid="stFileUploader"] section > div {
            gap: 10px !important; justify-content: center;
        }
        [data-testid="stFileUploader"] button {
            margin-left: 30px !important; 
        }
        [data-testid="stFileUploader"] ul { display: none !important; }
        [data-testid="stFileUploader"] div[role="progressbar"] { display: none !important; }
        [data-testid="stFileUploader"] label {
            font-size: 13px !important; color: #2563EB !important; margin-bottom: 0px !important;
        }

        /* [파일 정보 박스] */
        .uploaded-info-box {
            background-color: #ECFDF5;
            border: 1px solid #10B981;
            border-radius: 8px;
            height: 40px; 
            display: flex; align-items: center;
            padding-left: 10px; padding-right: 10px;
            color: #059669; font-weight: 600; font-size: 14px;
            margin-bottom: 4px;
        }
        
        /* [버튼 스타일] */
        .stButton > button, 
        .stDownloadButton > button {
            min-height: 38px !important; height: 38px !important;
            padding-top: 0px !important; padding-bottom: 0px !important;
            font-size: 15px !important;
        }
        [data-testid="stSidebar"] .stButton > button {
            height: 60px !important; font-size: 22px !important;
        }

        /* [데이터 표] */
        [data-testid="stDataFrame"] { font-size: 16px !important; }
        [data-testid="stDataFrameResizable"] div[role="columnheader"] {
            justify-content: center !important; text-align: center !important;
            font-weight: 700 !important; background-color: #F9FAFB;
        }
        [data-testid="stDataFrameResizable"] div[role="gridcell"] {
            justify-content: center !important; text-align: center !important;
        }

        /* [탭 스타일] */
        button[data-baseweb="tab"] {
            font-size: 16px !important;
            font-weight: 700 !important;
            color: #6B7280 !important;
            padding-top: 0px !important;
            padding-bottom: 0px !important;
            height: 50px !important;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            color: #2563EB !important;
            border-bottom-color: #2563EB !important;
        }
        div[data-testid="stTabs"] {
            gap: 0px !important;
        }
        .stTabs [data-baseweb="tab-panel"] {
            padding-top: 10px !important; 
        }
        </style>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 4. 페이지 함수
# -----------------------------------------------------------------------------
def page_home():
    col_left, col_right = st.columns([5, 5])

    # 결과 초기화 
    if 'result' not in st.session_state:
        st.session_state['result'] = None
        st.session_state['human'] = None
        st.session_state['group'] = None
        st.session_state['error_log'] = None
    
    # -------------------------------------------------------------------------
    # [좌측 패널]
    # -------------------------------------------------------------------------
    with col_left:
        with st.container():            
            # 헤더
            h_col1, h_col2 = st.columns([7.5, 2.5], gap="small")
            with h_col1:
                st.markdown('<div class="card-title" style="margin-top: 5px;">📂 인력 배치 프로그램 </div>', unsafe_allow_html=True)
            with h_col2:
                try:
                    with open("template/template.xlsx", "rb") as file:
                        st.download_button(
                            label="📄 조건 양식 다운로드",
                            data=file,
                            file_name="template.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                except FileNotFoundError:
                    st.error("양식 파일이 없습니다.")

            st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

            # 업로더
            u_col1, u_col2 = st.columns([7, 3], gap="small")
            
            with u_col1:
                uploaded_file = st.file_uploader(
                    "파일 선택", 
                    type=['xlsx'], 
                    label_visibility="collapsed",
                    accept_multiple_files=False,
                    key=f"uploader_{st.session_state['uploader_key']}" 
                )
            
            with u_col2:
                if uploaded_file:
                    st.markdown(f'''
                        <div class="uploaded-info-box">
                            <span style="margin-right:6px;">✅</span>
                            <span style="flex-grow:1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                                {uploaded_file.name}
                            </span>
                        </div>
                    ''', unsafe_allow_html=True)
                    if st.button("🗑️ 파일 제거", type="secondary", use_container_width=True):
                        reset_uploader()
                        st.rerun()
                else:
                    st.markdown('''
                        <div style="
                            background-color: #F3F4F6; border: 1px dashed #D1D5DB; border-radius: 8px;
                            height: 80px; display: flex; align-items: center; justify-content: center;
                            color: #9CA3AF; font-size: 13px;">
                            파일 정보 대기 중
                        </div>
                    ''', unsafe_allow_html=True)

            # 데이터 로드
            if uploaded_file:
                try:
                    df_raw = pd.read_excel(uploaded_file)
                    if df_raw.shape[1] >= 7:
                        workers = int(df_raw.iloc[1,7])
                        raw_df = df_raw.iloc[1:,0:7].fillna(0)
                        raw_df.columns = ['구분','진료과그룹','근무지','인력_Min','인력_Max','월별_Min','월별_Max']
                    else:
                        raw_df = pd.DataFrame(columns=['구분','진료과그룹','근무지','인력_Min','인력_Max','월별_Min','월별_Max'])
                except:
                     raw_df = pd.DataFrame(columns=['구분','진료과그룹','근무지','인력_Min','인력_Max','월별_Min','월별_Max'])
            else:
                raw_df = pd.DataFrame(columns=['구분','진료과그룹','근무지','인력_Min','인력_Max','월별_Min','월별_Max'])
            
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            
            # [좌측] 높이 700px
            df =st.data_editor(
                raw_df, 
                use_container_width=True, 
                hide_index=True, 
                num_rows="dynamic", 
                height=700 
            )

    # -------------------------------------------------------------------------
    # [우측 패널] Action & Analysis
    # -------------------------------------------------------------------------
    with col_right:
        with st.container():
            # 헤더: 버튼들이 잘리지 않도록 비율 조정 (5:5)
            rh_col1, rh_col2 = st.columns([5, 5], gap="medium")
            with rh_col1:
                title_html = '<div class="card-title" style="margin-top: 5px; display: flex; align-items: center; gap: 10px;">'
                title_html += '<span style="white-space: nowrap;">🚀 Action & Analysis</span>'
                
                # 에러 로그가 있으면 옆에 표시 (공간이 좁으므로 최대 너비 제한)
                if st.session_state.get('error_log'):
                    error_msg = st.session_state['error_log']
                    title_html += f'''
                        <div style="
                            background-color: #FEE2E2; 
                            color: #B91C1C; 
                            border: 1px solid #FECACA;
                            border-radius: 6px; 
                            padding: 2px 8px; 
                            font-size: 13px; 
                            font-weight: 600;
                            white-space: nowrap;
                            overflow: hidden;
                            text-overflow: ellipsis;
                            max-width: 250px;
                            cursor: help;
                        " title="{error_msg}">
                             ⚠️ {error_msg}
                        </div>
                    '''
                title_html += '</div>'
                st.markdown(title_html, unsafe_allow_html=True)
            with rh_col2:
                # 버튼들 간의 간격을 위해 컬럼 비율 조정
                col1, col2 = st.columns([1, 1], gap="small")
                with col1:
                    if st.button("⚡ 최적화 실행", type="primary", use_container_width=True, disabled=df.empty):
                        with st.spinner("데이터 분석 중..."):
                            try:
                                final = WORKFORCE_ASSIGN(df=df,workers=workers,n=3)
                                final.modeling()
                                
                                if final.result is not None:
                                    st.session_state['result'] = final.result.reset_index() # 결과 데이터 프레임 생성 및 상태 저장 
                                    st.session_state['human'] = final.worker_counts.reset_index()
                                    st.session_state['group'] = final.dept_counts_by_month.reset_index()
                                    st.session_state['error_log'] = None
                                else:
                                    st.session_state['result'] = None
                                    st.session_state['human'] = None
                                    st.session_state['group'] = None
                                    st.session_state['error_log'] = getattr(final, 'error_log', "알 수 없는 최적화 오류")
                            except Exception as e:
                                st.session_state['result'] = None
                                st.session_state['error_log'] = f"코드 실행 오류: {str(e)}"
                                st.error(f"실행 중 오류가 발생했습니다: {e}")
                with col2:
                    # 엑셀 다운로드 로직
                    if st.session_state.get('result') is not None and not st.session_state['result'].empty:
                        # 엑셀 파일 생성 함수 호출
                        excel_buffer = create_excel_file(
                            st.session_state['result'], 
                            st.session_state['human'], 
                            st.session_state['group'], 
                            df # 설정(Out Dept) 확인용
                        )
                        
                        if excel_buffer:
                            download_data = excel_buffer.getvalue()
                            st.download_button(
                                label="📜 Excel 다운",
                                data=download_data,
                                file_name="배정결과_통합.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                    else:
                        st.button('📜 Excel 다운', disabled=True, use_container_width=True)
            
            # 탭 구성
            tab1, tab2, tab3 = st.tabs(["📋 배정결과", "👥 인력별집계", "📊 구분별집계"])
            
            # Placeholder 함수
            def show_placeholder(icon, text, is_error=False):
                bg_color = "#FEF2F2" if is_error else "#F9FAFB"
                border_color = "#FECACA" if is_error else "#D1D5DB"
                text_color = "#B91C1C" if is_error else "#9CA3AF"
                
                st.markdown(f'''
                    <div style="
                        height: 750px; 
                        background-color:{bg_color}; 
                        border-radius:8px; 
                        display:flex; 
                        flex-direction:column; 
                        align-items:center; 
                        justify-content:center; 
                        color:{text_color}; 
                        border: 1px dashed {border_color};
                        text-align: center;
                        padding: 20px;
                    ">
                        <div style="font-size: 50px; margin-bottom: 10px;">{icon}</div>
                        <div style="font-size: 1.1rem; font-weight: 600;">{text}</div>
                    </div>
                ''', unsafe_allow_html=True)

            # -----------------------------------------------------------------
            # [Tab 1] 배정결과
            # -----------------------------------------------------------------
            with tab1:
                if st.session_state['result'] is None:
                    if st.session_state.get('error_log'):
                        show_placeholder("⚠️", f"최적화 실패<br><br><small>{st.session_state['error_log']}</small>", is_error=True)
                    else:
                        show_placeholder("👥", "최적화 실행 후<br><b>집계</b>가 표시됩니다.")                    
                else:
                    st.dataframe(
                        st.session_state['result'],
                        use_container_width=True, 
                        height=750,
                        hide_index=True
                    )

            with tab2:
                if st.session_state['result'] is None:
                    if st.session_state.get('error_log'):
                        show_placeholder("⚠️", "최적화 실패로 인한<br>데이터 없음", is_error=True)
                    else:
                        show_placeholder("👥", "최적화 실행 후<br><b>인력별 집계</b>가 표시됩니다.")
                else:
                    st.dataframe(
                        st.session_state['human'],
                        use_container_width=True, 
                        height=750,
                        hide_index=True
                    )                
            with tab3:
                if st.session_state['result'] is None:
                    if st.session_state.get('error_log'):
                        show_placeholder("⚠️", "최적화 실패로 인한<br>데이터 없음", is_error=True)
                    else:
                        show_placeholder("👥", "최적화 실행 후<br><b>구분별 집계</b>가 표시됩니다.")
                else:
                    st.dataframe(
                        st.session_state['group'],
                        use_container_width=True, 
                        height=750,
                        hide_index=True
                    )

def main():
    set_dashboard_style()
    page_home()


if __name__ == "__main__":
    main()