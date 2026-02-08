import pandas as pd
import io
import os

def create_excel_file(result, human_df, group_df, df):
    if result is None or result.empty:
        return None

    # 1. 설정 정보 추출 (Out Departments 식별 - 근무지 기준)
    # df 컬럼: ['구분','진료과그룹','근무지','인력_Min','인력_Max','월별_Min','월별_Max']
    out_depts = []
    if not df.empty and '구분' in df.columns and '근무지' in df.columns:
        for idx, row in df.iterrows():
            dept_name = row['구분']
            location = str(row['근무지']).lower()
            # 근무지에 'out'이 포함되면 파견으로 간주
            if 'out' in location:
                out_depts.append(dept_name)
    
    # 2. Excel 버퍼 생성
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # 3. 데이터프레임 시트 작성 (Sheet1)
    result.to_excel(writer, sheet_name='Sheet1')
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # 4. 서식(Format) 정의
    # (1) 헤더 서식: 하늘색 배경 + 일반 선 + 굵게
    header_fmt = workbook.add_format({
        'bg_color': '#CCEEFF',
        'border': 1,
        'bold': True,
        'align': 'center',
        'valign': 'vcenter'
    })

    # (2) 일반 셀 서식: 일반 선
    common_fmt = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # (3) 파견 셀 서식: 노란색 배경 + 일반 선
    dispatch_fmt = workbook.add_format({
        'bg_color': '#FFFF00',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # (4) 제 1열(인턴명) 서식: 오른쪽만 굵은 선 (두께 2)
    first_col_fmt = workbook.add_format({
        'left': 1, 'top': 1, 'bottom': 1, 
        'right': 2, # 우측 굵은 선
        'align': 'center',
        'valign': 'vcenter'
    })

    # 5. 디자인 적용 (데이터 루프)
    # 열 머리글(헤더) 쓰기
    for col_num, value in enumerate(result.columns.values):
        worksheet.write(0, col_num + 1, value, header_fmt)
        worksheet.write(0, 0, "Intern_ID", header_fmt) # 인덱스 헤더

    # 데이터 영역 쓰기
    # result.iterrows()는 (index, Series)를 반환
    # index: 인턴 이름
    for row_num, (index, row) in enumerate(result.iterrows()):
        # 첫 번째 열 (인턴 이름) - 우측 굵은 선 적용
        worksheet.write(row_num + 1, 0, index, first_col_fmt)
    
        # 나머지 데이터 셀
        for col_num, value in enumerate(row):
            # '파견' 단어 포함 여부 체크 (또는 out_depts에 포함 여부)
            # 여기서는 value가 진료과 이름이라고 가정
            if str(value) in out_depts:
                target_fmt = dispatch_fmt
            else:
                target_fmt = common_fmt
            
            worksheet.write(row_num + 1, col_num + 1, value, target_fmt)

    # ---------------------------------------------------------
    # 두 번째 표 (human_df -> intern_per_sum_df) 붙이기
    # ---------------------------------------------------------
    # 시작할 열 위치 계산: 첫 번째 표 열 개수(13개? result.columns) + 여유공간(2칸) = 15번째 열
    start_col = len(result.columns) + 2 

    # (1) 요약 표 헤더 쓰기
    for col_num, col_name in enumerate(human_df.columns):
        worksheet.write(0, start_col + col_num + 1, col_name, header_fmt)
    worksheet.write(0, start_col, "인턴 요약", header_fmt)

    # (2) 요약 표 데이터 쓰기
    for row_num, (index, row) in enumerate(human_df.iterrows()):
        # 요약 표의 첫 열 (인턴 ID 등)
        worksheet.write(row_num + 1, start_col, index, first_col_fmt)
        
        # 요약 표의 나머지 데이터
        for col_num, value in enumerate(row):
            # 데이터가 숫자나 일반 텍스트이므로 common_fmt 적용
            worksheet.write(row_num + 1, start_col + col_num + 1, value, common_fmt)

    # ---------------------------------------------------------
    # 세 번째 표 (group_df -> month_per_sum_df) 아래에 붙이기
    # ---------------------------------------------------------
    # 시작할 행 위치 계산: 첫 번째 표 행 개수(len(result)) + 헤더(1) + 여유공간(3칸) = 56행부터 시작
    start_row_3 = len(result) + 4 
    
    # (1) 세 번째 표 헤더 쓰기 (start_row_3 위치에)
    for col_num, col_name in enumerate(group_df.columns):
        worksheet.write(start_row_3, col_num + 1, col_name, header_fmt)
    worksheet.write(start_row_3, 0, "월별/과별 요약", header_fmt)

    # (2) 세 번째 표 데이터 쓰기
    for row_num, (index, row) in enumerate(group_df.iterrows()):
        # 실제 엑셀 행 번호: 시작 행 + 헤더(1) + 현재 루프 번호
        current_excel_row = start_row_3 + row_num + 1
    
        # 세 번째 표의 첫 열 (진료과명)
        worksheet.write(current_excel_row, 0, index, first_col_fmt)
        
        # 세 번째 표의 나머지 데이터
        for col_num, value in enumerate(row):
            worksheet.write(current_excel_row, col_num + 1, value, common_fmt)

    # 열 너비 정리
    worksheet.set_column(0, 0, 15)            # 메인 인턴ID
    worksheet.set_column(1, len(result.columns), 12) # 메인 월별 데이터
    worksheet.set_column(start_col, start_col, 15)   # 요약 표 첫 열
    worksheet.set_column(start_col + 1, start_col + 5, 12) # 요약 표 데이터

    writer.close()
    output.seek(0)
    return output