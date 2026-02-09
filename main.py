import sys
import subprocess
import os
import time
import random
import re
from datetime import datetime

# ==========================================
# 0. 라이브러리 자동 설치 및 임포트
# ==========================================
def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        print(f"📦 '{package}' 라이브러리가 없습니다. 자동으로 설치합니다...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

required_packages = ['pandas', 'openpyxl', 'rich']
for package in required_packages:
    install_and_import(package)

import pandas as pd
import numpy as np
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from rich.theme import Theme
from rich.status import Status
from rich.prompt import Prompt
from rich import box
from rich.align import Align

custom_theme = Theme({
    "info": "bold cyan",
    "warning": "bold yellow",
    "error": "bold red",
    "success": "bold green",
    "highlight": "bold magenta",
    "header": "bold white on dark_green",
})
console = Console(theme=custom_theme)

# ==========================================
# 1. 스마트 데이터 처리 함수
# ==========================================

def smart_get_year(value):
    try:
        if hasattr(value, 'year'): return value.year
        val_str = str(value).strip()
        if not val_str or val_str.lower() == 'nan': return 0
        if len(val_str) == 4 and val_str.isdigit(): return int(val_str)
        if '-' in val_str or len(val_str) == 6:
            prefix = val_str.split('-')[0] if '-' in val_str else val_str[:6]
            yy = int(prefix[:2])
            return 2000 + yy if 0 <= yy <= 30 else 1900 + yy
        return int(float(val_str))
    except: return 0 

def format_phone_number(value):
    try:
        s = str(value).strip()
        if not s or s.lower() == 'nan': return ""
        if s.endswith('.0'): s = s[:-2]
        s = s.replace('-', '').replace(' ', '').replace('.', '')
        if len(s) == 10 and s.startswith('1'): s = '0' + s
        if len(s) == 11: return f"{s[:3]}-{s[3:7]}-{s[7:]}"
        return s
    except: return value

def format_gender_output(value):
    s = str(value).strip()
    if s in ['남자', '남', 'Male', 'M', 'Man']: return '남자'
    elif s in ['여자', '여', 'Female', 'F', 'Woman']: return '여자'
    return s

def normalize_gender(value):
    s = str(value)
    for char in ['\ufeff', '\u200b', '\xa0', ' ', '\t', '\n', '?']:
        s = s.replace(char, '')
    s = s.strip()
    if '남' in s or 'Man' in s or 'Male' in s or 'M' in s: return '남'
    if '여' in s or 'Woman' in s or 'Female' in s or 'F' in s: return '여'
    return s

def get_name_key(name):
    name = str(name).strip()
    if len(name) > 1: return name[1:] 
    return name

# [UPGRADE] 특수문자(+), 다중 공백 처리 강화된 분리 함수
def split_school_and_name(raw_text):
    text = str(raw_text).strip()
    
    # 1. 4글자 이하는 이름으로 간주
    if len(text) <= 4:
        return None, text

    # 2. 전처리: 구분자로 쓰일만한 특수문자(+, /, ,, _)를 공백으로 치환
    # 예: "한국고+홍길동" -> "한국고 홍길동"
    cleaned_text = re.sub(r'[\+\,\/\_\-]', ' ', text)
    
    # 3. 전처리: 여러 개의 공백을 하나의 공백으로 축소
    # 예: "한국고    홍길동" -> "한국고 홍길동"
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()

    # 4. 공백 기준 분리 시도
    if ' ' in cleaned_text:
        parts = cleaned_text.split()
        if len(parts) >= 2:
            potential_school = parts[0]
            # 학교 접미사 확인
            if any(potential_school.endswith(suffix) for suffix in ['고', '학교']):
                real_name = " ".join(parts[1:]) # 나머지는 이름
                return potential_school, real_name

    # 5. 공백 없이 붙어있는 경우 (기존 정규식 활용)
    # 예: "한국고홍길동"
    match = re.search(r"^(.*?)(고등학교|학교|고)([가-힣]{2,4})$", text)
    if match:
        school_part = match.group(1) + match.group(2)
        name_part = match.group(3)
        if not match.group(1): 
            return None, text
        return school_part, name_part

    return None, text

def calculate_score(group_id, member, group_status, constraints, weights, ignore_age=False, limits_config=None):
    leader_min_year = constraints['leader_years'].get(group_id, 0)
    
    if limits_config:
        if group_status[group_id]['count'] >= limits_config['total']: return -float('inf')
        my_gender = member['gender']
        current_gender_cnt = group_status[group_id]['genders'].get(my_gender, 0)
        if my_gender == '남' and current_gender_cnt >= limits_config['male']: return -float('inf')
        elif my_gender == '여' and current_gender_cnt >= limits_config['female']: return -float('inf')

    if not ignore_age and leader_min_year > 0 and member['birth_year'] < leader_min_year:
        return -float('inf')

    if member['name_key'] in group_status[group_id]['names']:
        return -float('inf')

    nc_count = group_status[group_id]['new_cam'].get(member['new_cam'], 0)
    if nc_count >= 3:
        return -float('inf')

    score = 0
    score -= group_status[group_id]['count'] * weights['size'] 
    score -= group_status[group_id]['genders'].get(member['gender'], 0) * weights['gender']
    score -= group_status[group_id]['majors'].get(member['major'], 0) * weights['major']
    score -= group_status[group_id]['birth_years'].get(member['birth_year'], 0) * weights['birth_year']
    
    if nc_count == 1: score += weights['new_cam_cluster_bonus'] 
    elif nc_count > 0: score += weights['new_cam_exist_bonus']
    else: score -= weights['new_cam_scatter_penalty']

    return score

# ==========================================
# 2. 메인 로직 클래스
# ==========================================

class TeamBuilder:
    def __init__(self):
        self.num_groups = 0
        self.leaders = {} 
        self.members = []
        self.df_members = None
        self.result_groups = {} 
        self.original_name_col = None 
        
        self.weights = {
            'size': 100,            
            'gender': 50,           
            'major': 40,            
            'birth_year': 30,       
            'new_cam_cluster_bonus': 200,   
            'new_cam_exist_bonus': 20,      
            'new_cam_scatter_penalty': 50,  
            'new_cam_max_penalty': 100      
        }

    def agree_to_terms(self):
        print()
        title_text = Text("🦄 GIST 새내기배움터 조 자동 배정 프로그램 🚀", style="bold white", justify="center")
        subtitle_text = Text("\nFairness • Balance • Optimization", style="dim white", justify="center")
        dept_text = Text("\n[ 지스트 문화행사위원회 ] [ v 26 . 1 . 21 ]", style="bold bright_green", justify="center")
        header_content = Text.assemble(title_text, subtitle_text, dept_text, justify="center")
        console.print(Panel(header_content, box=box.DOUBLE, border_style="bright_green", padding=(1, 4), style="on black"))
        print("\n")

        terms_text = (
            "[bold green]1. 개인정보 처리 안내[/bold green]\n"
            "   본 프로그램은 로컬 환경(내 컴퓨터)에서 엑셀 데이터를 읽고 처리할 뿐, 외부 서버로 전송하지 않습니다.\n"
            "   다만, 개인정보가 포함된 파일을 다룰 때는 유출되지 않도록 각별히 주의해주세요.\n\n"
            "[bold green]2. 면책 조항[/bold green]\n"
            "   해당 프로그램은 조 배정을 돕기 위한 보조 도구이며, 최종 책임은 사용자에게 있습니다.\n\n"
            "[bold green]3. 사용 목적[/bold green]\n"
            "   본 프로그램은 새내기배움터 및 학과 행사를 위해서만 사용하여야 합니다."
        )

        console.print(Panel(terms_text, title="[bold white]📜 이용 약관 및 안내 사항[/bold white]", border_style="green", padding=(1, 2)))
        
        while True:
            choice = console.input("\n[bold yellow]위 사항에 모두 동의하십니까? (y/n): [/bold yellow]").strip().lower()
            if choice == 'y':
                console.print("\n[dim]약관에 동의하셨습니다. 데이터 입력을 시작합니다...[/dim]\n")
                return True
            elif choice == 'n':
                console.print("\n[bold red]동의하지 않으셨으므로 프로그램을 종료합니다.[/bold red]")
                return False
            else:
                console.print("[red]⚠️ 'y' 또는 'n'만 입력해주세요.[/red]")

    def load_data(self):
        year_input = console.input("[bold yellow]⚡ 행사 연도 (4자리)를 입력하세요[/bold yellow] [dim](예: 2025)[/dim]: ").strip()
        leader_file = f"{year_input}ST_leader.xlsx"
        member_file = f"{year_input}ST_freshmen.xlsx"
        
        console.print(f"\n[bold]📂 파일 탐색 중...[/bold]")
        if not os.path.exists(leader_file):
            console.print(f"[error]❌ 오류: '{leader_file}' 파일을 찾을 수 없습니다.[/error]")
            return False
        if not os.path.exists(member_file):
            console.print(f"[error]❌ 오류: '{member_file}' 파일을 찾을 수 없습니다.[/error]")
            return False

        console.print(f"   [success]✔[/success] 조장 파일: [underline]{leader_file}[/underline]")
        console.print(f"   [success]✔[/success] 참가자 파일: [underline]{member_file}[/underline]")

        try:
            with console.status("[bold cyan]조장 정보를 분석하는 중...", spinner="dots"):
                l_df = pd.read_excel(leader_file)
                year_cols = [c for c in l_df.columns if '생년' in str(c) or 'year' in str(c).lower()]
                group_col = [c for c in l_df.columns if '조' in str(c) and '번호' in str(c)]
                
                use_indices = (len(year_cols) < 2)
                max_group_id = 0

                for idx, row in l_df.iterrows():
                    try:
                        g_id = int(row[group_col[0]]) if group_col else int(row.iloc[0])
                        if g_id > max_group_id: max_group_id = g_id
                        
                        if use_indices:
                            y1, y2 = smart_get_year(row.iloc[2]), smart_get_year(row.iloc[4])
                        else:
                            y1, y2 = smart_get_year(row[year_cols[0]]), smart_get_year(row[year_cols[1]])
                        self.leaders[g_id] = min(y1, y2)
                    except: continue
                
                self.num_groups = max_group_id
            
            if self.num_groups == 0:
                console.print("[error]❌ 조 번호를 인식하지 못했습니다.[/error]")
                return False
            console.print(f"   [info]➜ 총 조 개수:[/info] [highlight]{self.num_groups}개[/highlight]")

        except Exception as e:
            console.print(f"[error]❌ 조장 파일 읽기 실패: {e}[/error]")
            return False

        try:
            with console.status("[bold cyan]참가자 데이터를 처리하는 중...", spinner="dots"):
                m_df = pd.read_excel(member_file)
                
                cols = m_df.columns
                col_map = {}
                for c in cols:
                    c_str = str(c)
                    if '성명' in c_str or '이름' in c_str: col_map['name'] = c
                    elif '고교' in c_str or '고등학교' in c_str: col_map['highschool'] = c
                    elif '신캠' in c_str: col_map['new_cam'] = c
                    elif '학과' in c_str or '학부' in c_str: col_map['major'] = c
                    elif '성별' in c_str: col_map['gender'] = c
                    elif '전화' in c_str or '휴대폰' in c_str: col_map['phone'] = c
                    elif '생년' in c_str or '주민' in c_str: col_map['birth'] = c

                if 'name' in col_map:
                    self.original_name_col = col_map['name']

                required_keys = ['name', 'new_cam', 'major', 'gender', 'phone', 'birth']
                missing_keys = [k for k in required_keys if k not in col_map]

                if missing_keys:
                    console.print(f"[error]❌ 필수 컬럼 찾기 실패.[/error]")
                    console.print(f"[dim]누락된 항목: {missing_keys}[/dim]")
                    return False

                processed_members = []
                for idx, row in m_df.iterrows():
                    raw_name_input = str(row[col_map['name']])
                    
                    extracted_school, real_name = split_school_and_name(raw_name_input)
                    
                    excel_school = ""
                    if 'highschool' in col_map:
                        val = str(row[col_map['highschool']])
                        if val and val.lower() != 'nan':
                            excel_school = val
                    
                    final_highschool = excel_school if excel_school else (extracted_school if extracted_school else "")

                    raw_birth = row[col_map['birth']]
                    raw_gender = row[col_map['gender']]
                    gender = normalize_gender(raw_gender)
                    
                    info = {
                        'original_idx': idx,
                        'name': real_name, 
                        'name_key': get_name_key(real_name),
                        'birth_year': smart_get_year(raw_birth),
                        'gender': gender, 
                        'major': row[col_map['major']],
                        'new_cam': int(row[col_map['new_cam']]),
                        'highschool': final_highschool, 
                        'phone': row[col_map['phone']],
                        'raw_birth': raw_birth
                    }
                    processed_members.append(info)
                
                self.members = processed_members
                self.df_members = m_df
            
            console.print(f"   [info]➜ 참가자 인원:[/info] [highlight]{len(self.members)}명[/highlight]")
            
        except Exception as e:
            console.print(f"[error]❌ 참가자 파일 읽기 실패: {e}[/error]")
            return False
            
        return True

    def assign_teams(self):
        total_m = sum(1 for m in self.members if m['gender'] == '남')
        total_f = sum(1 for m in self.members if m['gender'] == '여')
        
        base_m = total_m // self.num_groups
        rem_m = total_m % self.num_groups
        male_slots = [base_m + 1] * rem_m + [base_m] * (self.num_groups - rem_m)
        
        base_f = total_f // self.num_groups
        rem_f = total_f % self.num_groups
        female_slots = [base_f + 1] * rem_f + [base_f] * (self.num_groups - rem_f)
        
        max_retries = 2000 
        success = False
        
        with console.status("[bold green]성비와 인원을 완벽하게 맞추는 중...[/bold green]", spinner="bouncingBar") as status:
            for attempt in range(1, max_retries + 1):
                random.shuffle(male_slots)
                random.shuffle(female_slots)
                
                temp_totals = [m + f for m, f in zip(male_slots, female_slots)]
                if max(temp_totals) - min(temp_totals) > 1:
                    continue 

                self.male_limits = {i: c for i, c in zip(range(1, self.num_groups + 1), male_slots)}
                self.female_limits = {i: c for i, c in zip(range(1, self.num_groups + 1), female_slots)}
                self.total_limits = {i: m+f for i, m, f in zip(range(1, self.num_groups+1), male_slots, female_slots)}

                group_status = {
                    i: {
                        'count': 0, 'names': [], 'genders': {},
                        'majors': {}, 'birth_years': {}, 'new_cam': {}
                    } for i in range(1, self.num_groups + 1)
                }
                assignments = {m['original_idx']: None for m in self.members}
                
                sorted_members = sorted(self.members, key=lambda x: x['birth_year'])

                def try_assign(member_list, ignore_age=False):
                    failed = []
                    for member in member_list:
                        best_group = -1
                        best_score = -float('inf')
                        candidates = list(range(1, self.num_groups + 1))
                        random.shuffle(candidates)
                        
                        for g_id in candidates:
                            score = calculate_score(g_id, member, group_status, 
                                                 {'leader_years': self.leaders}, self.weights, 
                                                 ignore_age=ignore_age, 
                                                 limits_config={
                                                     'total': self.total_limits[g_id],
                                                     'male': self.male_limits[g_id],
                                                     'female': self.female_limits[g_id]
                                                 })
                            if score > best_score:
                                best_score = score
                                best_group = g_id
                        
                        if best_group != -1 and best_score > -float('inf'):
                            assignments[member['original_idx']] = best_group
                            self._update_status(group_status, best_group, member)
                        else:
                            failed.append(member)
                    return failed

                unassigned = try_assign(sorted_members, ignore_age=False)
                if unassigned: unassigned = try_assign(unassigned, ignore_age=False)
                if unassigned: unassigned = try_assign(unassigned, ignore_age=True)
                if unassigned: unassigned = try_assign(unassigned, ignore_age=True)

                if not unassigned:
                    success = True
                    self.result_groups = assignments
                    break
                
                if attempt % 100 == 0:
                    status.update(f"[bold yellow]재시도 중... (Attempt {attempt})[/bold yellow]")

        if success:
            console.print(f"\n[success]✨ 배정 성공! (총 시도: {attempt}회)[/success]")
            self._print_stats(group_status)
        else:
            console.print(f"\n[error]❌ [치명적 오류] {max_retries}번을 시도했으나 배정에 실패했습니다.[/error]")
            console.print("이유: 동명이인 등 하드 조건이 너무 까다롭거나 성비 슬롯 매칭이 어렵습니다.")

    def _update_status(self, status, g_id, member):
        st = status[g_id]
        st['count'] += 1
        st['names'].append(member['name_key'])
        g = member['gender']
        st['genders'][g] = st['genders'].get(g, 0) + 1
        m = member['major']
        st['majors'][m] = st['majors'].get(m, 0) + 1
        b = member['birth_year']
        st['birth_years'][b] = st['birth_years'].get(b, 0) + 1
        nc = member['new_cam']
        st['new_cam'][nc] = st['new_cam'].get(nc, 0) + 1

    def _print_stats(self, status):
        table = Table(title="📊 [bold]최종 배정 결과[/bold]", border_style="cyan", header_style="bold white on dark_green")
        table.add_column("조 이름", justify="center", style="bold cyan")
        table.add_column("인원", justify="center", style="white")
        table.add_column("성비 (남/여)", justify="center")
        table.add_column("비고", justify="left")

        for g_id in range(1, self.num_groups + 1):
            st = status[g_id]
            m_cnt = st['genders'].get('남', 0)
            f_cnt = st['genders'].get('여', 0)
            gender_str = f"[cyan]{m_cnt}[/cyan] : [magenta]{f_cnt}[/magenta]"
            table.add_row(f"새터 {g_id}조", f"{st['count']}명", gender_str, "배정 완료")
        console.print(table)

    def save_result(self):
        console.print("\n[bold]💾 결과 저장 중...[/bold]")
        
        self.df_members['최종 배정 조'] = self.df_members.index.map(self.result_groups)
        
        # 이름과 고등학교 정보를 새로운 컬럼에 할당
        clean_names = {}
        clean_schools = {}
        
        for m in self.members:
            clean_names[m['original_idx']] = m['name']
            clean_schools[m['original_idx']] = m['highschool']

        self.df_members['성명'] = self.df_members.index.map(clean_names)
        self.df_members['출신고교명'] = self.df_members.index.map(clean_schools)

        # 포맷팅
        phone_col = None
        for c in self.df_members.columns:
            if '전화' in str(c) or 'phone' in str(c).lower():
                phone_col = c
                break
        if phone_col:
            self.df_members[phone_col] = self.df_members[phone_col].apply(format_phone_number)

        gender_col = next((c for c in self.df_members.columns if '성별' in str(c) or 'gender' in str(c).lower()), None)
        if gender_col:
            self.df_members[gender_col] = self.df_members[gender_col].apply(format_gender_output)

        # 4. 컬럼 순서 정리 및 [중복 컬럼 제거]
        cols = self.df_members.columns.tolist()
        target_order = ['최종 배정 조', '성명', '출신고교명'] 
        
        priority_cols = ['신캠조', '학과', '학부', '성별', '전화번호', '생년월일']
        added = set(target_order)
        
        for p_key in priority_cols:
            for c in cols:
                if p_key in c and c not in added:
                    target_order.append(c)
                    added.add(c)
                    break
                    
        for c in cols:
            # 1. 이미 추가된 컬럼이면 건너뜀
            if c in added: continue
            # 2. 원본 이름 컬럼(출신고교명+이름) 이면 건너뜀
            if c == self.original_name_col: continue 
            
            target_order.append(c)
        
        final_df = self.df_members[target_order].sort_values(by=['최종 배정 조', '성명'])
        
        filename = f"team_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        final_df.to_excel(filename, index=False)
        
        console.print(f"[success]✔ 모든 작업이 완료되었습니다![/success]")
        console.print(f"   📂 저장된 파일: [underline bold]{filename}[/underline bold]\n")

if __name__ == "__main__":
    builder = TeamBuilder()
    if builder.agree_to_terms():
        if builder.load_data():
            builder.assign_teams()
            builder.save_result()