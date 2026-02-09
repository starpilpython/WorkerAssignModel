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
        self.constraints_list = [] # 제약조건 저장 리스트
        self.error_log = None # [신규] 최적화 실패 원인 저장
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
        # 제약함수 수집
        #----------------------------------        
        
        ## (제약조건 1) 근무인원은 무조건 월별 1곳 배치
        for e in self.employees_index:
            for m in self.months:
                self.constraints_list.append((pulp.lpSum([x[e][m][d] for d in self.departments]) == 1, f"Assignment_1Dept_Per_Month_{e}_{m}"))
        

        ## (제약조건 2) 월별로 배치된 진료과 당 인턴 수
        for d in self.departments:
            for m in self.months:
                self.constraints_list.append((pulp.lpSum([x[e][m][d] for e in self.employees_index]) >= self.dept_config[d]['limit_m'][0], f"Dept_Capacity_Min_{d}_{m}"))
                self.constraints_list.append((pulp.lpSum([x[e][m][d] for e in self.employees_index]) <= self.dept_config[d]['limit_m'][-1], f"Dept_Capacity_Max_{d}_{m}"))
        

        ## (제약조건 3) 인력별 부서 할당 횟수 (그룹화로 처리)
        department_group_map = defaultdict(list) 
        for dept, info in self.dept_config.items():
            group_name = info['department_group'][0]
            key = dept if group_name == 'A' else group_name
            department_group_map[key].append(dept)

        for e in self.employees_index:
            for group_key, d_list in department_group_map.items():
                min_i = 0
                max_i = 0
                for dept in d_list:
                    min_i += self.dept_config[dept]['limit_i'][0]
                    max_i += self.dept_config[dept]['limit_i'][-1]
                self.constraints_list.append((pulp.lpSum([x[e][m][d] for m in self.months for d in d_list]) >= min_i, f"Worker_Group_Min_{e}_{group_key}"))
                self.constraints_list.append((pulp.lpSum([x[e][m][d] for m in self.months for d in d_list]) <= max_i, f"Worker_Group_Max_{e}_{group_key}"))
        
        
        ## (제약조건 4) 파견병원은 최대 파견병원 횟수 제한
        out_departments = [d for d, info in self.dept_config.items() if info['location_group'][0].startswith('out')]
        for e in self.employees_index:
            self.constraints_list.append((pulp.lpSum([x[e][m][d] for m in self.months for d in out_departments]) <= self.out_group_count, f"Global_Out_Max_{e}"))
            self.constraints_list.append((pulp.lpSum([x[e][m][d] for m in self.months for d in out_departments]) >= self.out_group_count - 2, f"Global_Out_Min_{e}"))
      
      
        ## (제약조건 5) 연속 근무 및 장소 그룹 제약
        all_locations = set(info['location_group'][0] for info in self.dept_config.values())
        for e in self.employees_index:
            for loc in all_locations:
                if loc == 'out1': continue
                d_list = [d for d, info in self.dept_config.items() if info['location_group'][0] == loc]
                if loc == 'main':
                    for d in d_list:
                        for m_idx in range(len(self.months) - 1):
                            m1, m2 = self.months[m_idx], self.months[m_idx + 1]
                            self.constraints_list.append((x[e][m1][d] + x[e][m2][d] <= 1, f"No_Cont_Dept_{e}_{d}_{m1}"))
                else:
                    for m_idx in range(len(self.months) - 1):
                        m1, m2 = self.months[m_idx], self.months[m_idx + 1]
                        self.constraints_list.append((pulp.lpSum([x[e][m1][d] for d in d_list]) + pulp.lpSum([x[e][m2][d] for d in d_list]) <= 1, f"No_Cont_Loc_{e}_{loc}_{m1}"))

        ## (제약조건 6) out1 강제 연속 근무 및 배타적 파견
        all_out_depts = [d for d, info in self.dept_config.items() if info['location_group'][0].startswith('out')]
        out1_depts = [d for d, info in self.dept_config.items() if info['location_group'][0] == 'out1']
        y = pulp.LpVariable.dicts("y_start", (self.employees_index, range(len(self.months)-1)), cat='Binary')

        for e in self.employees_index:
            self.constraints_list.append((pulp.lpSum([y[e][m] for m in range(len(self.months)-1)]) <= 1, f"Out1_Start_MaxOnce_{e}"))
            for m in range(len(self.months) - 1):
                m1, m2 = self.months[m], self.months[m+1]
                self.constraints_list.append((pulp.lpSum([x[e][m1][d] for d in out1_depts]) >= y[e][m], f"Out1_ForcedM1_{e}_{m}"))
                self.constraints_list.append((pulp.lpSum([x[e][m2][d] for d in out1_depts]) >= y[e][m], f"Out1_ForcedM2_{e}_{m}"))
                for d in out1_depts:
                    self.constraints_list.append((x[e][m1][d] + x[e][m2][d] <= 2 - y[e][m], f"Out1_CrossRule_{e}_{d}_{m}"))
                other_months = [month for month in self.months if month not in [m1, m2]]
                self.constraints_list.append((pulp.lpSum([x[e][om][d] for om in other_months for d in all_out_depts]) <= 100 * (1 - y[e][m]), f"Out1_Exclusion_OtherOuts_{e}_{m}"))

        for m in range(len(self.months) - 1):
            self.constraints_list.append((pulp.lpSum([y[e][m] for e in self.employees_index]) == 1, f"Out1_Monthly_StarterCount_{m}"))

        #----------------------------------
        # 초기 제약조건 적용 및 실행
        #----------------------------------
        for ct, name in self.constraints_list:
            prob += ct, name

        prob.writeLP("intern_debug.lp")
        prob.solve() 
        print(f'[DEBUG] 분석상태: {pulp.LpStatus[prob.status]} (code: {prob.status})')

        # 1. 성공한 경우 (Optimal)
        if pulp.LpStatus[prob.status] == 'Optimal':
            result_data = []
            for m in self.months:
                for e in self.employees_index: 
                    for d in self.departments:
                        val = pulp.value(x[e][m][d])
                        if val is not None and round(val) == 1:
                            result_data.append({'Month': m, 'Employee': e, 'Dept': d})
            
            if result_data:
                self.result = pd.DataFrame(result_data).pivot(index='Employee', columns='Month', values='Dept')
                self.result.index = self.employees_index
                self.result.columns = self.months
                self._short()
                self.error_log = None
            else:
                self.result = None
                self.error_log = "최적해를 찾았으나 배정 데이터가 생성되지 않았습니다 (모델 설정 오류)."

        # 2. 불능인 경우 (Infeasible) -> 진단 루프 실행
        elif pulp.LpStatus[prob.status] == 'Infeasible':
            self.result = None
            self._run_diagnostic()

        # 3. 기타 오류 (Undefined, Not Solved 등)
        else:
            self.result = None
            self.error_log = f"최적화 실패: {pulp.LpStatus[prob.status]} (데이터가 너무 복잡하거나 제약이 너무 많습니다.)"
            print(f"[ERROR] {self.error_log}")

    def _run_diagnostic(self):
        print("\n" + "="*50)
        print("[CRITICAL] 최적화 불능(Infeasible) 발생. 원인 분석을 시작합니다...")
        print(f"총 제약조건 {len(self.constraints_list)}개를 대상으로 이진 탐색을 수행합니다.")
        print("="*50)
        
        low = 0
        high = len(self.constraints_list) - 1
        culprit_idx = -1

        while low <= high:
            mid = (low + high) // 2
            test_prob = pulp.LpProblem("Infeasible_Analysis", pulp.LpMinimize)
            test_prob += 0
            for i in range(mid + 1):
                ct, name = self.constraints_list[i]
                test_prob += ct, name
            
            test_prob.solve(pulp.PULP_CBC_CMD(msg=0))
            
            if pulp.LpStatus[test_prob.status] == 'Infeasible':
                culprit_idx = mid
                high = mid - 1
            else:
                low = mid + 1

        if culprit_idx != -1:
            culprit_name = self.constraints_list[culprit_idx][1]
            self.error_log = f"충돌 규칙: {culprit_name}"
            print(f"\n[발견] 원인 제약조건: {culprit_name}")
        else:
            self.error_log = "제약조건 간의 복합적인 충돌로 특정 원인을 찾을 수 없습니다."
        
        print("="*50 + "\n")
        
                
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