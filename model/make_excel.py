import pandas as pd
import io

def create_excel_file(result, human_df, group_df, df):
    if result is None or result.empty:
        return None

    # 1. 설정 정보 추출 (Out Departments 식별)
    out_depts = []
    if not df.empty and '구분' in df.columns and '근무지' in df.columns:
        for idx, row in df.iterrows():
            dept_name = row['구분']
            location = str(row['근무지']).lower()
            if 'out' in location:
                out_depts.append(dept_name)
    
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # 데이터프레임 시트 생성 (실제 데이터 쓰기는 루프에서 수동으로 진행하므로 빈 시트 확보용)
    result.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # 2. 서식 정의
    header_fmt = workbook.add_format({'bg_color': '#CCEEFF', 'border': 1, 'bold': True, 'align': 'center', 'valign': 'vcenter'})
    common_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    dispatch_fmt = workbook.add_format({'bg_color': '#FFFF00', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    # 첫 열(데이터의 실제 첫 번째 값) 강조용 서식
    first_col_fmt = workbook.add_format({'left': 1, 'top': 1, 'bottom': 1, 'right': 2, 'align': 'center', 'valign': 'vcenter'})

    # ---------------------------------------------------------
    # 3. 메인 표 (첫 번째 표) - 인덱스 제거 및 데이터 쓰기
    # ---------------------------------------------------------
    # 헤더 쓰기 (인덱스 열 없이 바로 시작)
    for col_num, value in enumerate(result.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    # 데이터 쓰기
    for row_num, (idx, row) in enumerate(result.iterrows()):
        for col_num, value in enumerate(row):
            # 첫 번째 열 데이터(Worker_ID 등)에는 굵은 우측 선 적용
            if col_num == 0:
                target_fmt = first_col_fmt
            elif str(value) in out_depts:
                target_fmt = dispatch_fmt
            else:
                target_fmt = common_fmt
            
            worksheet.write(row_num + 1, col_num, value, target_fmt)

    # ---------------------------------------------------------
    # 4. 두 번째 표 (인턴 요약) - 인덱스 제거 및 우측 배치
    # ---------------------------------------------------------
    # 시작 열 계산 (첫 표 끝난 후 1칸 여유)
    start_col = len(result.columns) + 1 

    # 헤더 쓰기
    for col_num, col_name in enumerate(human_df.columns):
        worksheet.write(0, start_col + col_num, col_name, header_fmt)

    # 데이터 쓰기
    for row_num, (idx, row) in enumerate(human_df.iterrows()):
        for col_num, value in enumerate(row):
            # 두 번째 표의 첫 열 서식 적용
            target_fmt = first_col_fmt if col_num == 0 else common_fmt
            worksheet.write(row_num + 1, start_col + col_num, value, target_fmt)

    # ---------------------------------------------------------
    # 5. 세 번째 표 (월별/과별 요약) - 인덱스 제거 및 하단 배치
    # ---------------------------------------------------------
    # 시작 행 계산 (첫 표 끝난 후 3칸 여유)
    start_row_3 = len(result) + 3 
    
    # 헤더 쓰기
    for col_num, col_name in enumerate(group_df.columns):
        worksheet.write(start_row_3, col_num, col_name, header_fmt)

    # 데이터 쓰기
    for row_num, (idx, row) in enumerate(group_df.iterrows()):
        current_excel_row = start_row_3 + row_num + 1
        for col_num, value in enumerate(row):
            target_fmt = first_col_fmt if col_num == 0 else common_fmt
            worksheet.write(current_excel_row, col_num, value, target_fmt)

    # 6. 열 너비 정리
    worksheet.set_column(0, 0, 15) # 메인 표 첫 열
    worksheet.set_column(1, len(result.columns) - 1, 12) # 메인 데이터
    worksheet.set_column(start_col, start_col, 15) # 두 번째 표 첫 열
    worksheet.set_column(start_col + 1, start_col + len(human_df.columns), 12)

    writer.close()
    output.seek(0)
    return output