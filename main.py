import sys
import subprocess
import os
import time
import random
import re
import unicodedata # [NEW] 글자 너비 계산을 위해 추가
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
# 1. 스마트 데이터 처리 및 검증 함수
# ==========================================

# [NEW] 화면상 실제 너비를 계산하는 함수 (한글=2, 영어/숫자=1)
def get_display_width(s):
    width = 0
    for char in s:
        if unicodedata.east_asian_width(char) in ('F', 'W', 'A'):
            width += 2
        else:
            width += 1
    return width

# [FIXED] 너비 기반 정렬로 2글자 이름도 줄 맞춤 완벽 해결
def format_list_2col(items):
    if not items: return ""
    lines = []
    target_width = 40 # 1열의 목표 너비 (넉넉하게 설정)
    
    for i in range(0, len(items), 2):
        item1 = items[i]
        item2 = items[i+1] if i+1 < len(items) else ""
        
        # 실제 차지하는 너비 계산
        w1 = get_display_width(item1)
        # 필요한 공백 개수 계산
        padding = max(1, target_width - w1)
        
        lines.append(f"{item1}{' ' * padding}{item2}")
        
    return "\n".join(lines)

def normalize_id_str(value):
    s = str(value).strip()
    if not s or s.lower() == 'nan': return ""
    if s.endswith('.0'): s = s[:-2]
    s = re.sub(r'[\s\-]', '', s)
    if len(s) == 12 and s.isdigit():
        s = '0' + s
    return s

def smart_get_year(value):
    try:
        if hasattr(value, 'year'): return value.year
        val_str = str(value).strip()
        if not val_str or val_str.lower() == 'nan': return 0
        
        normalized_id = normalize_id_str(val_str)
        
        if len(normalized_id) == 13 and normalized_id.isdigit():
            yy = int(normalized_id[:2])
            gender_digit = int(normalized_id[6])
            if gender_digit in [3, 4, 7, 8]:
                return 2000 + yy
            else:
                return 1900 + yy
        
        if len(normalized_id) == 6 and normalized_id.isdigit():
             yy = int(normalized_id[:2])
             if 0 <= yy <= 30: return 2000 + yy
             else: return 1900 + yy

        if len(val_str) == 4 and val_str.isdigit():
            return int(val_str)
            
        return int(float(val_str))
    except: return 0 

def format_phone_number(value):
    try:
        if isinstance(value, float):
            value = int(value)
        s = str(value).strip()
        if not s or s.lower() == 'nan': return ""
        if s.endswith('.0'): s = s[:-2]
        s = s.replace('-', '').replace(' ', '').replace('.', '')
        if len(s) == 10 and s.startswith('1'): s = '0' + s
        if len(s) == 11: return f"{s[:3]}-{s[3:7]}-{s[7:]}"
        return s
    except: return value

def format_id_number(value):
    s = normalize_id_str(value)
    if len(s) == 13 and s.isdigit():
        return f"{s[:6]}-{s[6:]}"
    return value

def validate_id_checksum(id_val):
    s = normalize_id_str(id_val)
    if len(s) != 13 or not s.isdigit():
        return False
    weights = [2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5]
    total = sum(int(s[i]) * weights[i] for i in range(12))
    check_digit = (11 - (total % 11)) % 10
    return check_digit == int(s[12])

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

def normalize_school_name(text):
    text = str(text).strip()
    text = text.replace('고등학교', '고')
    text = text.replace('과고', '과학고')
    return text

def normalize_dept_name(text):
    text = str(text).strip()
    if '도전탐색과정' in text:
        return '도전탐색과정'
    return text

def split_school_and_name(raw_text):
    text = str(raw_text).strip()
    if len(text) <= 4: return None, text

    cleaned_text = re.sub(r'[\+\,\/\_\-]', ' ', text)
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()

    # 1. 공백이 있는 경우: 맨 뒤 어절이 이름인지 확인 (Right-to-Left Strategy)
    if ' ' in cleaned_text:
        parts = cleaned_text.split()
        last_part = parts[-1]
        
        # 맨 뒤가 2~4글자 한글이면 이름으로 간주
        if 2 <= len(last_part) <= 4 and re.match(r'^[가-힣]+$', last_part):
            real_name = last_part
            # 나머지 앞부분을 전부 합쳐서 학교명으로 (공백 제거)
            potential_school = "".join(parts[:-1]) 
            return potential_school, real_name

    # 2. 공백이 없거나 위 조건이 안 맞는 경우 정규식 사용
    match = re.search(r"^(.*?)(고등학교|학교|고)\s?([가-힣]{2,4})$", text)
    if match:
        school_part = match.group(1) + match.group(2)
        name_part = match.group(3)
        school_part = school_part.replace(" ", "")
        if not match.group(1): return None, text
        return school_part, name_part
        
    return None, text

# [UPGRADE] 절대 균등 분배(Max Diff <= 1) 알고리즘
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

    score = 0
    score -= group_status[group_id]['count'] * weights['size'] 
    score -= group_status[group_id]['genders'].get(member['gender'], 0) * weights['gender']
    score -= group_status[group_id]['majors'].get(member['major'], 0) * weights['major']
    score -= group_status[group_id]['birth_years'].get(member['birth_year'], 0) * weights['birth_year']

    nc = member['new_cam']
    nc_count = group_status[group_id]['new_cam'].get(nc, 0)
    
    nc_base = constraints.get('new_cam_bases', {}).get(nc, 0) 
    nc_rem = constraints.get('new_cam_rems', {}).get(nc, 0)   
    
    if nc_count >= nc_base + 1:
        return -float('inf')
        
    if nc_count == nc_base:
        groups_at_max = sum(1 for g in group_status.values() if g['new_cam'].get(nc, 0) > nc_base)
        if groups_at_max >= nc_rem:
            return -float('inf')
            
    if nc_count < nc_base:
        score += 1000000 
    elif nc_count == nc_base:
        score += 50 

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
        }

    def agree_to_terms(self):
        print()
        title_text = Text("🦄 GIST 새내기배움터 조 자동 배정 프로그램 🚀", style="bold white", justify="center")
        subtitle_text = Text("\nFairness • Balance • Optimization", style="dim white", justify="center")
        dept_text = Text("\n[ 지스트 문화행사위원회 ] [ v 26 . 2 . 13 ]", style="bold bright_green", justify="center")
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
        year_input_str = console.input("[bold yellow]⚡ 행사 연도 (4자리)를 입력하세요[/bold yellow] [dim](예: 2026)[/dim]: ").strip()
        try:
            event_year = int(year_input_str)
        except:
            event_year = 2026
            
        leader_file = f"{year_input_str}ST_leader.xlsx"
        member_file = f"{year_input_str}ST_freshmen.xlsx"
        
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
            
        except Exception as e:
            console.print(f"[error]❌ 조장 파일 읽기 실패: {e}[/error]")
            return False

        try:
            with console.status("[bold cyan]참가자 데이터를 처리 및 검증하는 중...", spinner="dots"):
                m_df = pd.read_excel(member_file)
                
                cols = m_df.columns
                col_map = {}
                
                # [SIMPLE] 전화번호 컬럼 매핑: 휴대폰/전화 보이면 바로 끝!
                for c in cols:
                    c_str = str(c).strip()
                    if 'phone' not in col_map and ('휴대폰' in c_str or '전화' in c_str): 
                        col_map['phone'] = c
                        break # 하나 찾으면 바로 루프 종료 (First-Win)
                
                # 나머지 컬럼 매핑
                for c in cols:
                    c_str = str(c).strip() 
                    if 'name' not in col_map and ('성명' in c_str or '이름' in c_str): col_map['name'] = c
                    elif 'highschool' not in col_map and ('고교' in c_str or '고등학교' in c_str): col_map['highschool'] = c
                    elif 'new_cam' not in col_map and '신캠' in c_str: col_map['new_cam'] = c
                    elif 'major' not in col_map and ('학과' in c_str or '학부' in c_str): col_map['major'] = c
                    elif 'gender' not in col_map and '성별' in c_str: col_map['gender'] = c
                    elif 'birth' not in col_map and ('생년' in c_str or '주민' in c_str): col_map['birth'] = c

                if 'name' in col_map:
                    self.original_name_col = col_map['name']

                required_keys = ['name', 'new_cam', 'major', 'gender', 'phone', 'birth']
                missing_keys = [k for k in required_keys if k not in col_map]

                if missing_keys:
                    console.print(f"[error]❌ 필수 컬럼 찾기 실패.[/error]")
                    console.print(f"[dim]누락된 항목: {missing_keys}[/dim]")
                    return False

                processed_members = []
                
                duplicates_log = []
                minors_log = []
                invalid_id_log = [] 
                missing_school_log = []
                
                seen_phones = set()
                seen_ids = set()

            # ------------------------------------------------------------------
            # [Step 2] 인터랙티브 데이터 검증 및 수정
            # ------------------------------------------------------------------
            
            print()
            console.print("[bold yellow]⚡ 데이터 검증 및 대화형 수정을 시작합니다...[/bold yellow]")
            time.sleep(0.5)

            for idx, row in m_df.iloc[::-1].iterrows():
                raw_name_input = str(row[col_map['name']])
                new_cam_val = int(row[col_map['new_cam']])
                
                extracted_school, real_name = split_school_and_name(raw_name_input)
                
                needs_correction = False
                issue_type = ""
                
                if not extracted_school and len(real_name) > 4:
                    needs_correction = True
                    issue_type = "학교명/이름 뭉침 의심"
                elif extracted_school and not any(x in extracted_school for x in ['고', '학교', 'High', 'Academy']):
                    needs_correction = True
                    issue_type = "학교명 형식 오류 의심"
                elif re.search(r'[\(\)0-9]', real_name) or len(real_name) > 5:
                    needs_correction = True
                    issue_type = "이름 형식 오류 (특수문자/길이)"

                elif not extracted_school and len(real_name) <= 4:
                    has_excel_school = False
                    if 'highschool' in col_map:
                        val = str(row[col_map['highschool']])
                        if val and val.lower() != 'nan' and val.strip():
                            has_excel_school = True
                    
                    if not has_excel_school:
                        needs_correction = True
                        issue_type = "학교명 미기재"

                if needs_correction:
                    console.print(Panel(
                        f"[bold red]⚠️  데이터 확인 필요 ({issue_type})[/bold red]\n"
                        f"입력값: [yellow]{raw_name_input}[/yellow]\n\n"
                        f"현재 인식된 학교: [cyan]{extracted_school if extracted_school else '없음'}[/cyan]\n"
                        f"현재 인식된 이름: [cyan]{real_name}[/cyan]",
                        border_style="red"
                    ))
                    
                    user_resp = Prompt.ask(
                        "[bold white]수정하시겠습니까?[/bold white] (올바른 '[green]학교명 이름[/green]' 입력 / 엔터치면 원본 유지)",
                        default="",
                        show_default=False
                    ).strip()
                    
                    if user_resp:
                        new_school, new_name = split_school_and_name(user_resp)
                        if not new_school and len(new_name) <= 4:
                            extracted_school = None
                            real_name = new_name
                        else:
                            extracted_school = new_school
                            real_name = new_name
                        console.print(f"[green]✔ 수정됨: 학교='{extracted_school}', 이름='{real_name}'[/green]")
                    else:
                        console.print("[dim]원본 유지 (또는 미기재)[/dim]")
                    print()

                if not extracted_school and len(real_name) <= 4:
                    excel_school_val = ""
                    if 'highschool' in col_map:
                        val = str(row[col_map['highschool']])
                        if val and val.lower() != 'nan':
                            excel_school_val = val.strip()
                    
                    if not excel_school_val:
                        missing_school_log.append(f"{real_name} (신캠 {new_cam_val}조)")
                        extracted_school = "미기재" 

                excel_school = ""
                if 'highschool' in col_map:
                    val = str(row[col_map['highschool']])
                    if val and val.lower() != 'nan':
                        excel_school = val
                
                final_highschool = excel_school if excel_school else (extracted_school if extracted_school else "미기재")
                final_highschool = normalize_school_name(final_highschool)

                clean_phone = format_phone_number(row[col_map['phone']])
                raw_birth = row[col_map['birth']]
                clean_id = normalize_id_str(raw_birth) 
                birth_year = smart_get_year(raw_birth)
                clean_major = normalize_dept_name(row[col_map['major']])

                is_duplicate = False
                if clean_phone and clean_phone in seen_phones: is_duplicate = True
                if clean_id and len(clean_id) == 13 and clean_id in seen_ids: is_duplicate = True
                
                if is_duplicate:
                    duplicates_log.append(f"{real_name} (신캠 {new_cam_val}조)")
                    continue 
                
                if clean_phone: seen_phones.add(clean_phone)
                if clean_id and len(clean_id) == 13: seen_ids.add(clean_id)

                is_minor = birth_year >= (event_year - 18)
                is_valid_id = False
                if clean_id and len(clean_id) == 13:
                    is_valid_id = validate_id_checksum(clean_id)
                
                if is_minor:
                    if is_valid_id:
                        minors_log.append(f"{real_name} (신캠 {new_cam_val}조)")
                        continue
                    else:
                        invalid_id_log.append(f"{real_name} (신캠 {new_cam_val}조) - [bold red]미성년/오기재 불분명[/bold red]")
                
                elif not is_minor and (clean_id and len(clean_id) == 13) and not is_valid_id:
                    invalid_id_log.append(f"{real_name} (신캠 {new_cam_val}조)")

                raw_gender = row[col_map['gender']]
                gender = normalize_gender(raw_gender)
                
                info = {
                    'original_idx': idx,
                    'name': real_name, 
                    'name_key': get_name_key(real_name),
                    'birth_year': birth_year,
                    'gender': gender, 
                    'major': clean_major,
                    'new_cam': new_cam_val,
                    'highschool': final_highschool, 
                    'phone': row[col_map['phone']], 
                    'raw_birth': raw_birth,         
                    'clean_id': clean_id 
                }
                processed_members.append(info)
            
            self.members = processed_members[::-1]
            self.df_members = m_df
            
            if duplicates_log or minors_log or invalid_id_log or missing_school_log:
                print()
                console.print("[bold white on dark_red] 🚨 데이터 검증 및 전처리 리포트 [/bold white on dark_red]", justify="left")
                time.sleep(0.5)
                
                if duplicates_log:
                    console.print(Panel(format_list_2col(duplicates_log), title="[bold red]중복 제출자 자동 삭제됨 (최신 데이터 유지)[/bold red]", border_style="red"))
                    time.sleep(0.3)
                
                if minors_log:
                    console.print(Panel(format_list_2col(minors_log), title="[bold red]미성년자 지원자 자동 삭제됨[/bold red]", border_style="red"))
                    time.sleep(0.3)

                if invalid_id_log:
                    console.print(Panel(format_list_2col(invalid_id_log), title="[bold yellow]주민번호 오기재 확인 (확인 필요 - 데이터 유지됨)[/bold yellow]", border_style="yellow"))
                    time.sleep(0.3)
                    
                if missing_school_log:
                    console.print(Panel(format_list_2col(missing_school_log), title="[bold yellow]학교명 미기재 (확인 필요 - '미기재'로 저장됨)[/bold yellow]", border_style="yellow"))
                    time.sleep(0.3)

            print()
            console.print("[bold white on dark_green] ✅ 데이터 처리 완료 [/bold white on dark_green]", justify="left")
            time.sleep(0.5)
            console.print(f"   [info]➜ 총 조 개수:[/info] [highlight]{self.num_groups}개[/highlight]")
            console.print(f"   [info]➜ 최종 참가자 인원:[/info] [highlight]{len(self.members)}명[/highlight]")
            time.sleep(0.5)
            
        except Exception as e:
            console.print(f"[error]❌ 참가자 파일 읽기 실패: {e}[/error]")
            import traceback
            traceback.print_exc()
            return False
            
        return True

    def assign_teams(self):
        print()
        time.sleep(0.5)

        total_m = sum(1 for m in self.members if m['gender'] == '남')
        total_f = sum(1 for m in self.members if m['gender'] == '여')
        
        base_m = total_m // self.num_groups
        rem_m = total_m % self.num_groups
        male_slots = [base_m + 1] * rem_m + [base_m] * (self.num_groups - rem_m)
        
        base_f = total_f // self.num_groups
        rem_f = total_f % self.num_groups
        female_slots = [base_f + 1] * rem_f + [base_f] * (self.num_groups - rem_f)

        new_cam_counts = {}
        for m in self.members:
            nc = m['new_cam']
            new_cam_counts[nc] = new_cam_counts.get(nc, 0) + 1
            
        new_cam_bases = {}
        new_cam_rems = {}
        for nc, count in new_cam_counts.items():
            new_cam_bases[nc] = count // self.num_groups  
            new_cam_rems[nc] = count % self.num_groups    
        
        max_retries = 10000 
        success = False
        
        with console.status("[bold green]성비와 인원을 완벽하게 맞추는 중... (엄격한 균등 분배)[/bold green]", spinner="bouncingBar") as status:
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
                                                 {'leader_years': self.leaders, 
                                                  'new_cam_bases': new_cam_bases,
                                                  'new_cam_rems': new_cam_rems}, 
                                                 self.weights, 
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
                    status.update(f"[bold yellow]수학적 최적화 탐색 중... (Attempt {attempt})[/bold yellow]")

        if success:
            console.print(f"\n[success]✨ 배정 성공! (총 시도: {attempt}회)[/success]")
            self._print_stats(group_status)
        else:
            console.print(f"\n[error]❌ [치명적 오류] {max_retries}번을 시도했으나 배정에 실패했습니다.[/error]")
            console.print("이유: 하드 조건(성비, 신캠조 엄격 분배, 동명이인)을 모두 만족하는 조합을 찾지 못했습니다.")

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
        print()
        time.sleep(0.5) 

        table = Table(title="📊 [bold]최종 배정 결과[/bold]", border_style="cyan", header_style="bold white on dark_green")
        table.add_column("조 이름", justify="center", style="bold cyan")
        table.add_column("인원", justify="center", style="white")
        table.add_column("성비\n(남/여)", justify="center")
        
        all_new_cams = set()
        for g_id in range(1, self.num_groups + 1):
            all_new_cams.update(status[g_id]['new_cam'].keys())
        sorted_new_cams = sorted(list(all_new_cams))
        
        for nc in sorted_new_cams:
            table.add_column(f"신캠\n{nc}조", justify="center")

        for g_id in range(1, self.num_groups + 1):
            st = status[g_id]
            m_cnt = st['genders'].get('남', 0)
            f_cnt = st['genders'].get('여', 0)
            gender_str = f"[cyan]{m_cnt}[/cyan] : [magenta]{f_cnt}[/magenta]"
            
            row_data = [f"새터 {g_id}조", f"{st['count']}명", gender_str]
            
            for nc in sorted_new_cams:
                cnt = st['new_cam'].get(nc, 0)
                
                if cnt == 0:
                    cnt_str = "[dim]0[/dim]"  
                elif cnt == 1:
                    cnt_str = f"[bold yellow]{cnt}[/bold yellow]"  
                elif cnt >= 3:
                    cnt_str = f"[bold red]{cnt}[/bold red]"  
                else: 
                    cnt_str = f"[bold green]{cnt}[/bold green]" 
                    
                row_data.append(cnt_str)
                
            table.add_row(*row_data)
            
        console.print(table)

    def save_result(self):
        print()
        time.sleep(0.5) 
        console.print("\n[bold]💾 결과 저장 중...[/bold]")
        
        valid_indices = [m['original_idx'] for m in self.members]
        final_df = self.df_members.loc[valid_indices].copy()
        
        final_df['최종 배정 조'] = final_df.index.map(self.result_groups)
        
        clean_names = {m['original_idx']: m['name'] for m in self.members}
        clean_schools = {m['original_idx']: m['highschool'] for m in self.members}
        clean_majors = {m['original_idx']: m['major'] for m in self.members}
        
        final_df['성명'] = final_df.index.map(clean_names)
        final_df['출신고교명'] = final_df.index.map(clean_schools)
        
        major_col = next((c for c in final_df.columns if '학과' in str(c) or '학부' in str(c)), None)
        if major_col:
            final_df[major_col] = final_df.index.map(clean_majors)

        phone_col = None
        for c in final_df.columns:
            if '전화' in str(c) or 'phone' in str(c).lower():
                phone_col = c
                break
        if phone_col:
            final_df[phone_col] = final_df[phone_col].apply(format_phone_number)

        id_col = next((c for c in final_df.columns if '주민' in str(c) or 'birth' in str(c).lower()), None)
        if id_col:
            final_df[id_col] = final_df[id_col].apply(format_id_number)

        gender_col = next((c for c in final_df.columns if '성별' in str(c) or 'gender' in str(c).lower()), None)
        if gender_col:
            final_df[gender_col] = final_df[gender_col].apply(format_gender_output)

        cols = final_df.columns.tolist()
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
            if c in added: continue
            if c == self.original_name_col: continue 
            target_order.append(c)
        
        final_df = final_df[target_order].sort_values(by=['최종 배정 조', '성명'])
        
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