'''
해당 파일은 매년 인턴 인력 배정에 도움이 되고자 만든 프로그램 
모델링 패키지 pulp (정수계획법으로 적용)
스트림릿으로 배포 예정
'''

# --------------------------------------------
# 패키지 로드 
import pandas as pd
import pulp
import os
import sys
from collections import defaultdict

# --------------------------------------------
# 클래스 설정

class WORKFORCE_ASSIGN:
    
    '''초기 실행'''
    def __init__(self,df,workers,n):
        self.df = df # 데이터프레임 설정
        self.workers = workers
        self.out_group_count = n # 파견병원 총 제한 횟수
        self.continue_work = ['out1'] # 연속 근무 허용
        self._setting()

    '''설정 실행'''
    def _setting(self):
        #----------------------------------
        # (판다스 경고) 미래의 동작 방식을 명시적으로 수용
        #----------------------------------
        pd.set_option('future.no_silent_downcasting', True)


        #----------------------------------
        # 제약조건 및 기타 필요사항 로드 
        #----------------------------------
        ## 데이터프레임 -> 딕셔러로 변환
        self.dept_config = {} 
        for idx, d in enumerate(self.df['구분']):
            dummy = self.df.iloc[idx]
            # 공통 데이터 먼저 생성
            config = {
                'department_group': [dummy['진료과그룹']],
                'location_group':[dummy['근무지']],
                'limit_i': [dummy['인력_Min'], dummy['인력_Max']],
                'limit_m': [dummy['월별_Min'], dummy['월별_Max']],
            }
            self.dept_config[d] = config
      
        ## 진료과 부서 설정
        self.departments = list(self.dept_config.keys())

        ## 근무인력 정리
        self.employees_index = [f'Worker_{x+1}' for x in range(self.workers)] 

        ## 월 설정
        self.months = [f'{x+1}월' for x in range(12)] 
        print('[DEBUG] 조건 파일 로드 완료')

    '''모델링설정'''
    def modeling(self):
        #----------------------------------
        # 정수계획법 setting
        #----------------------------------
        prob = pulp.LpProblem("Intern_Scheduling_Joker_Enabled", pulp.LpMinimize)
        prob += 0  # 상수 목적함수
        x = pulp.LpVariable.dicts("x", (self.employees_index, self.months, self.departments), cat='Binary')

        #----------------------------------
        # 제약함수
        #----------------------------------        
        
        ## (제약조건 1) 근무인원은 무조건 월별 1곳 배치
        for e in self.employees_index:
            for m in self.months:
                prob += pulp.lpSum([x[e][m][d] for d in self.departments]) == 1
        

        ## (제약조건 2) 월별로 배치된 진료과 당 인턴 수
        for d in self.departments:
            for m in self.months:
                prob += pulp.lpSum([x[e][m][d] for e in self.employees_index]) >= self.dept_config[d]['limit_m'][0]
                prob += pulp.lpSum([x[e][m][d] for e in self.employees_index]) <= self.dept_config[d]['limit_m'][-1]        
        

        ## (제약조건 3) 인력별 부서 할당 횟수 (그룹화로 처리)
        ### 데이터를 담을 그릇 생성
        department_group_map = defaultdict(list) 

        ### dept_config를 순회하며 분류
        for dept, info in self.dept_config.items():
            # info['group']은 ['A'] 형태의 리스트이므로 첫 번째 요소를 추출
            group_name = info['department_group'][0]
            
            # 핵심 전략: 'A'그룹은 과 이름을, 나머지는 그룹명을 Key로 사용
            key = dept if group_name == 'A' else group_name
            department_group_map[key].append(dept)

        ### 각 인력별 최소, 최대 횟수 설정
        for e in self.employees_index:
            for group_key, d_list in department_group_map.items():
                min_i = 0
                max_i = 0
                for dept in d_list:
                    min_i += self.dept_config[dept]['limit_i'][0]
                    max_i += self.dept_config[dept]['limit_i'][-1]
                    
                # 해당 그룹(d_list)에 속한 모든 과의 총 근무 개월 수 합산
                prob += pulp.lpSum([x[e][m][d] for m in self.months for d in d_list]) >= min_i
                prob += pulp.lpSum([x[e][m][d] for m in self.months for d in d_list]) <= max_i
        
        
        ## (제약조건 4) 파견병원은 최대 파견병원 횟수 ~ 파견병원 횟수회로 제한
        ### 모든 Out 계열 과 리스트 추출
        out_departments = [d for d, info in self.dept_config.items() if info['location_group'][0].startswith('out')]

        ### 제약조건 추가
        for e in self.employees_index:
            # 각 인턴별로 전체 기간(months) 동안 모든 out_departments에 배정된 횟수의 합
            # 최대값: out_group_count
            prob += pulp.lpSum([x[e][m][d] for m in self.months for d in out_departments]) <= self.out_group_count
            
            # 최소값: out_group_count - 1
            prob +=  pulp.lpSum([x[e][m][d] for m in self.months for d in out_departments]) >= self.out_group_count - 2        
      
      
        ## (제약조건 5) out1을 제외한 out2** 은 out 기준으로 연속 근무 제약 추가 / main인 경우에는 각 진료과별로 연속 근무 금지
        ### 모든 장소 그룹 추출
        all_locations = set(info['location_group'][0] for info in self.dept_config.values())

        ### out1을 제외한 그룹들에 대해 연속 근무 금지 제약 추가
        for e in self.employees_index:
            for loc in all_locations:
                # 1. out1은 강제 연속이므로 제외
                if loc == 'out1':
                    continue
                    
                d_list = [d for d, info in self.dept_config.items() if info['location_group'][0] == loc]
                
                # 2. Main 그룹인 경우: 각 진료과별로 '동일 과' 연속 근무만 금지
                if loc == 'main':
                    for d in d_list:
                        for m_idx in range(len(self.months) - 1):
                            m1, m2 = self.months[m_idx], self.months[m_idx + 1]
                            # 같은 직원이 같은 과(d)에 연속 2달 있을 수 없음
                            prob += x[e][m1][d] + x[e][m2][d] <= 1, f"No_Cont_Dept_{e}_{d}_{m1}"
                
                # 3. 그 외 파견지(out2, out3 등)인 경우: '장소 그룹' 전체 연속 근무 금지
                else:
                    for m_idx in range(len(self.months) - 1):
                        m1, m2 = self.months[m_idx], self.months[m_idx + 1]
                        # 같은 직원이 해당 파견지 그룹 내 어떤 과로든 연속 2달 있을 수 없음
                        prob += pulp.lpSum([x[e][m1][d] for d in d_list]) + \
                                pulp.lpSum([x[e][m2][d] for d in d_list]) <= 1, f"No_Cont_Loc_{e}_{loc}_{m1}"


        ## (제약조건 6) out1 2개월 파견시 타 병원 파견 불가 및 1인은 무조건 연속 2개월 근무
        ### 모든 out 계열 과 리스트 (out1, out2, out3 등 전체)
        all_out_depts = [d for d, info in self.dept_config.items() if info['location_group'][0].startswith('out')]
        
        ### 오직 out1(대우병원)에 속한 과 리스트
        out1_depts = [d for d, info in self.dept_config.items() if info['location_group'][0] == 'out1']

        ### y[e][m]: 직원 e가 m월에 out1 근무를 '시작'하면 1
        y = pulp.LpVariable.dicts("y_start", (self.employees_index, range(len(self.months)-1)), cat='Binary')

        for e in self.employees_index:
            # 한 직원이 수련 기간 중 out1 시작은 최대 한 번만 가능
            prob += pulp.lpSum([y[e][m] for m in range(len(self.months)-1)]) <= 1

            for m in range(len(self.months) - 1):
                m1, m2 = self.months[m], self.months[m+1]
                
                # [강제 연속] m월 시작 시 m, m+1월은 반드시 out1 근무
                prob += pulp.lpSum([x[e][m1][d] for d in out1_depts]) >= y[e][m]
                prob += pulp.lpSum([x[e][m2][d] for d in out1_depts]) >= y[e][m]
                
                # ===========================================================
                # [신규 추가] 세부 과 교차 근무 제약: m1과 m2에 같은 과 d를 중복해서 갈 수 없음
                for d in out1_depts:
                    prob += x[e][m1][d] + x[e][m2][d] <= 2 - y[e][m], f"Out1_Cross_{e}_{d}_{m1}"
                # ===========================================================

                # [핵심 수정: 전 기간 파견 금지]
                other_months = [month for month in self.months if month not in [m1, m2]]
                prob += pulp.lpSum([x[e][om][d] for om in other_months for d in all_out_depts]) <= 100 * (1 - y[e][m])

        for m in range(len(self.months) - 1):
            # m월에 out1 근무를 '새로 시작'하는 직원은 무조건 1명
            prob += pulp.lpSum([y[e][m] for e in self.employees_index]) == 1
 
        #----------------------------------
        # 최적해 구하기
        #----------------------------------   
        ## 계산
        prob.solve() 
        print(f'[DEBUG] 분석상태: {pulp.LpStatus[prob.status]}')
        
        ## 데이터 프레임으로 설정
        if pulp.LpStatus[prob.status] in ['Optimal', 'Not Solved']:
            result_data = []
            for m in self.months:
                for e in self.employees_index: 
                    for d in self.departments:
                        if pulp.value(x[e][m][d]) == 1:
                            dept_name = d
                            result_data.append({'Month': m, 'Employee': e, 'Dept': dept_name})
            self.result = pd.DataFrame(result_data).pivot(index='Employee', columns='Month', values='Dept')
            self.result.index = self.employees_index
            self.result.columns = self.months
            self._short() #집계표 출력
        else:
            self.worker_counts = None #근로자당 집계
            self.dept_counts_by_month = None #월별 집계
            self.result = None
        print(f"총 제약조건 개수: {len(prob.constraints)}")        
                
    '''집계표 출력'''
    def _short(self):
        # (근로자당 진료과 합계 생성)
        self.worker_counts = self.result.apply(lambda x: x.value_counts(), axis=1) \
              .reindex(columns=self.departments, fill_value=0) \
              .fillna(0).astype(int)
            
        # (월별 진료과 배치 합계 생성)
        self.dept_counts_by_month = self.result.apply(lambda x: x.value_counts()) \
                                    .reindex(index=self.departments, fill_value=0) \
                                    .fillna(0).astype(int)

# --------------------------------------------

if __name__ == '__main__':
    # 실행 파일(.exe)의 위치 파악
    if getattr(sys, 'frozen', False):
        current_path = os.path.dirname(sys.executable)
    else:
        current_path = os.path.dirname(os.path.abspath(__file__))

    # [수정] model 폴더 밖(상위 폴더)으로 한 단계 이동
    parent_path = os.path.abspath(os.path.join(current_path, ".."))

    # 상위 폴더에 있는 엑셀 파일 지정
    PATH_FILE = os.path.join(parent_path, "조건화면.xlsx")

    df = pd.read_excel(PATH_FILE)
    workers= int(df.iloc[1,7]) # 근무 인력 정리
    df = df.iloc[1:,0:7].copy()
    df.columns = ['구분','진료과그룹','근무지','인력_Min','인력_Max','월별_Min','월별_Max']
    df = df.fillna(0) # 칼럼 정리

    # 클래스 실행
    final = WORKFORCE_ASSIGN(df=df,workers=workers,n=3)
    final.modeling()