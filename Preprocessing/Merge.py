# *** 기능 ****************************************************************************************

# - 변수가이드에 맞춰서 플랫폼 활용 데이터 파일을 생성합니다.
# - NCSI 데이터 파일에서 변수가이드의 '변수레이블' 열만 가져온 후, LV/xLV열을 생성 및 정렬합니다.

# ************************************************************************************************
# *** 사용 방법 ***********************************************************************************

# 1. line 42~44의 경로를 알맞게 수정해주세요.
# 2. line 45에 분석 대상 연도를 알맞게 기입해주세요.(2자리로 기입해주세요.)

# ************************************************************************************************
# *** 주의사항 ************************************************************************************

# ── 백업 권장 ──
# • 실행 전 반드시 원본 파일을 별도 폴더에 백업해 두세요.
# • 문제가 발생하면 백업에서 복원 후, 파일명·열 이름·인코딩을 다시 확인하고 재실행합니다.

# - 코드에 사용되는 모든 파일은 닫혀있어야 합니다.

# NCSI_DATA_DIR:
# - 실제 디렉토리 경로가 정확히 존재해야 합니다.
# - 하위 파일명은 '(연도2자리)_NCSI_(업종명)_LV...(.xlsx/.xls/.csv)' 형식을 따라야합니다.

# GUIDE_DIR:
# - 디렉토리에 포함된 .xlsx 파일만 처리합니다.
# - 파일명에 '_(업종명)_' 패턴이 반드시 있어야합니다. (업종명은 NCSI_DATA_DIR에서 추출함)

# OUTPUT_DIR:
# - 존재하지 않을 경우 자동 생성되지만, 문제가 생긴다면 VS Code를 관리자 권한으로 실행해주세요.

# 변수가이드:
# - 변수가이드 파일 내 정확한 열 이름("변수레이블")이 일치해야 합니다.
# - LV/xLV 관계없이 NCSI 파일에 해당 열이 존재하는 경우 모델품질요인 칸에 작성하고,
# - 존재하지않는 경우, 비모델품질요인 칸에 작성되어야 합니다. (예시는 '병원' 샘플파일을 참고해주세요.)

# ************************************************************************************************

import os
import re
import pandas as pd
from collections import defaultdict
import sys

# *** 수정 부분 ***********************************************************************************
NCSI_DATA_DIR     = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Data"   
GUIDE_DIR     = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Guide"
OUTPUT_DIR    = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Result"
TARGET_YEARS  = ['24', '23', '22']
# ************************************************************************************************

VAR_LABEL_COL = "변수레이블"
LV_LABEL_GUIDE_COLUMNS = [4, 6]  # LV 레이블 (존재 필수) 열 인덱스
X_LV_LABEL_GUIDE_COLUMNS = [8, 10] # xLV 레이블 (존재 불가) 열 인덱스
# --- CONFIGURATION END ---

CODE_PATTERN = re.compile(r'^(x?LV\d+)$', re.IGNORECASE)

# --- 함수 정의 ---

def standardize_lv_name(col_name):
    """LV/xLV 컬럼 이름을 표준 형식('LVn', 'xLVn')으로 변환, 아니면 None 반환"""
    if not isinstance(col_name, str): return None
    col_name_stripped = col_name.strip() # 먼저 공백 제거
    match = CODE_PATTERN.match(col_name_stripped)
    if match:
        code_str = match.group(0); digits = ''.join(filter(str.isdigit, code_str))
        if not digits: return None
        if code_str.lower().startswith('xlv'): return 'xLV' + digits
        elif code_str.lower().startswith('lv'): return 'LV' + digits
    return None # 패턴 안맞으면 None

def maybe_standardize_name(name):
    """주어진 이름이 LV/xLV 패턴이면 표준화하고, 아니면 원본(strip된) 반환"""
    if not isinstance(name, str): return name # 문자열 아니면 그대로 반환
    name_stripped = name.strip()
    standardized = standardize_lv_name(name_stripped) # 표준화 시도
    return standardized if standardized is not None else name_stripped # 표준화 성공시 표준이름, 실패시 원본(strip된) 이름

def extract_industry_from_filename(filename):
    """
    파일명에서 '_industry_' 패턴을 찾아 업종명을 추출합니다.
    여러 패턴이 존재하면 가장 오른쪽에 있는 것을 반환합니다.
    패턴을 찾지 못하면 None을 반환합니다.
    """
    # 파일 확장자 제거
    name_without_ext, _ = os.path.splitext(filename)

    # '_'로 구분된 모든 부분을 찾음
    parts = name_without_ext.split('_')

    # '_'로 구분된 부분이 3개 이상이어야 '_something_' 패턴 존재 가능
    if len(parts) >= 3:
        # 마지막에서 두 번째 요소가 가장 오른쪽의 '_'와 '_' 사이의 내용
        industry_name = parts[-2]
        # 추출된 이름이 비어있지 않은지 확인
        if industry_name:
            return industry_name.strip() # 양쪽 공백 제거 후 반환
        else:
            return None # '_' 사이에 내용이 없으면 None 반환
    else:
        # '_something_' 패턴을 만들 수 없는 경우
        return None

def load_data_file(file_path):
    """단일 파일 로드 (CSV/Excel), 컬럼명 공백 제거"""
    _, ext = os.path.splitext(file_path.lower()); df = None
    try:
        if ext == '.csv':
            try: df = pd.read_csv(file_path, encoding='utf-8')
            except UnicodeDecodeError: print(f"[INFO] UTF-8 실패, CP949 시도: {os.path.basename(file_path)}"); df = pd.read_csv(file_path, encoding='cp949')
        elif ext in ('.xlsx', '.xls'): df = pd.read_excel(file_path)
        else: return None
        if df is not None:
            df.columns = [str(col).strip() for col in df.columns]
        return df
    except FileNotFoundError: print(f"[ERROR] 파일 없음: {file_path}"); return None
    except Exception as e:
        print(f"[WARN] 전체 파일 로드 오류 ({type(e).__name__}), 컬럼명만 읽기 시도: {file_path} -> {e}")
        try:
            if ext == '.csv':
                try: df_cols = pd.read_csv(file_path, encoding='utf-8', nrows=0)
                except UnicodeDecodeError: df_cols = pd.read_csv(file_path, encoding='cp949', nrows=0)
            elif ext in ('.xlsx', '.xls'): df_cols = pd.read_excel(file_path, nrows=0)
            else: return None
            df_cols.columns = [str(col).strip() for col in df_cols.columns]
            return list(df_cols.columns)
        except Exception as e_cols:
            print(f"[ERROR] 컬럼명 읽기 실패: {file_path} -> {e_cols}"); return None


# [수정됨] 가이드 레이블 로드 (LV/xLV만 표준화 시도 후 반환)
def get_validation_labels_from_guide(guide_path, lv_label_cols_indices, xlv_label_cols_indices):
    """가이드 파일에서 LV(존재 필수) 및 xLV(존재 불가) 레이블 목록 로드 (LV/xLV는 표준화)"""
    required_labels = set()
    forbidden_labels = set()
    try:
        guide_df = pd.read_excel(guide_path)
        guide_df.columns = [str(col).strip() for col in guide_df.columns]

        # LV 레이블 읽고 부분 표준화
        for col_idx in lv_label_cols_indices:
            if col_idx >= guide_df.shape[1]: continue
            try:
                 col_name = guide_df.columns[col_idx]
                 raw_labels = guide_df[col_name].dropna().astype(str).tolist() # strip은 maybe_standardize_name에서 처리
                 for label in raw_labels:
                     # LV/xLV면 표준화, 아니면 원본(strip된) 이름 추가
                     processed_label = maybe_standardize_name(label)
                     if processed_label: # 비어있지 않은 경우만 추가
                        required_labels.add(processed_label)
            except Exception as e:
                 print(f"[WARN] 가이드 '{os.path.basename(guide_path)}' LV 열 {col_idx} 처리 중 오류: {e}")

        # xLV 레이블 읽고 부분 표준화
        for col_idx in xlv_label_cols_indices:
             if col_idx >= guide_df.shape[1]: continue
             try:
                 col_name = guide_df.columns[col_idx]
                 raw_labels = guide_df[col_name].dropna().astype(str).tolist()
                 for label in raw_labels:
                     # LV/xLV면 표준화, 아니면 원본(strip된) 이름 추가
                     processed_label = maybe_standardize_name(label)
                     if processed_label:
                         forbidden_labels.add(processed_label)
             except Exception as e:
                 print(f"[WARN] 가이드 '{os.path.basename(guide_path)}' xLV 열 {col_idx} 처리 중 오류: {e}")

        print(f"  - 가이드({os.path.basename(guide_path)}): 검증용 필수 레이블 {len(required_labels)}개, 검증용 금지 레이블 {len(forbidden_labels)}개 로드됨 (LV/xLV 표준화됨).")
        # print(f"    - 검증 필수: {required_labels}") # 디버깅용
        # print(f"    - 검증 금지: {forbidden_labels}") # 디버깅용

        return required_labels, forbidden_labels

    except Exception as e:
        print(f"[ERROR] 가이드 로드 또는 검증 레이블 추출 실패: {guide_path} -> {e}", file=sys.stderr)
        return None, None

# [수정됨] 입력 파일 컬럼 로드 (LV/xLV만 표준화 시도 후 반환)
def get_combined_columns_from_files(file_paths):
    """주어진 파일 경로 리스트에서 모든 고유 컬럼명 집합 로드 (LV/xLV는 표준화)"""
    all_columns_raw = set()
    errors_loading = False
    for file_path in file_paths:
        loaded_content = load_data_file(file_path) # 컬럼명은 이미 strip됨
        if loaded_content is None:
            print(f"[WARN] 컬럼 검증 위해 파일 로드 실패: {os.path.basename(file_path)}")
            errors_loading = True
            continue

        if isinstance(loaded_content, pd.DataFrame):
            all_columns_raw.update(loaded_content.columns)
        elif isinstance(loaded_content, list):
             all_columns_raw.update(loaded_content)

    if errors_loading:
        print(f"[WARN] 일부 입력 파일 로드 실패로 컬럼 검증이 불완전할 수 있습니다.")

    # 부분 표준화 수행
    processed_columns = set()
    for col in all_columns_raw:
        # LV/xLV면 표준화, 아니면 원본(strip된) 이름 사용
        processed_col = maybe_standardize_name(col)
        if processed_col: # 비어있지 않은 경우만 추가
            processed_columns.add(processed_col)

    print(f"  - 입력 파일들에서 총 {len(processed_columns)}개의 고유 컬럼 확인됨 (LV/xLV 표준화됨).")
    # print(f"    - 확인된 컬럼 (부분 표준화): {processed_columns}") # 디버깅용
    return processed_columns


# --- 나머지 함수들은 이전 버전과 동일 ---

def load_lv_mapping(guide_path, var_col, code_cols_indices):
    """
    LV/xLV 매핑 및 우선순위 리스트 로드:
    - guide_df의 지정된 코드 열(code_cols_indices)에 있는 모든 코드에 대해
      각각 변수레이블(var_col)을 매핑
    - 우선순위 리스트는 코드 열 순서대로, 각 셀 출현 순으로 정렬
    """
    try:
        guide_df = pd.read_excel(guide_path)
        guide_df.columns = [str(col).strip() for col in guide_df.columns]
    except Exception as e:
        print(f"[ERROR] 가이드 로드 실패: {guide_path} -> {e}", file=sys.stderr)
        return None, None

    if var_col not in guide_df.columns:
        print(f"[ERROR] GUIDE '{var_col}' 없음: {guide_path}", file=sys.stderr)
        return None, None

    lv_mapping = {}
    processed_codes = set()
    # 1) 모든 코드→레이블 매핑 수집
    for _, row in guide_df.iterrows():
        raw_label = row[var_col]
        if pd.isna(raw_label):
            continue
        clean_label = str(raw_label).strip()

        for idx in code_cols_indices:
            if idx >= guide_df.shape[1]:
                continue
            raw_code = row.iloc[idx]
            if pd.notna(raw_code):
                code_str = str(raw_code).strip()
                if CODE_PATTERN.match(code_str):
                    # mapping
                    lv_mapping.setdefault(code_str, [])
                    if clean_label not in lv_mapping[code_str]:
                        lv_mapping[code_str].append(clean_label)
                    # 우선순위 리스트에 한 번만 추가
                    if code_str not in processed_codes:
                        processed_codes.add(code_str)

    # 2) 우선순위 리스트: 가이드 파일 코드 열 순서 & 행 순서대로 flatten
    codes_flat = []
    for idx in code_cols_indices:
        if idx >= guide_df.shape[1]:
            continue
        col_vals = guide_df.iloc[:, idx].dropna().astype(str).str.strip().tolist()
        codes_flat.extend(col_vals)

    codes_in_priority_order = []
    seen = set()
    for code in codes_flat:
        if CODE_PATTERN.match(code) and code in processed_codes and code not in seen:
            codes_in_priority_order.append(code)
            seen.add(code)

    return lv_mapping, codes_in_priority_order  # lv_mapping 키/리스트값, 우선순위 리스트 모두 원본 형태


def load_target_columns_from_guide(guide_path, var_col):
    """가이드 파일 '변수레이블' 열에서 최종 포함할 컬럼 목록 로드 (공백 제거)"""
    try:
        guide_df = pd.read_excel(guide_path)
        guide_df.columns = [str(col).strip() for col in guide_df.columns]
    except Exception as e: print(f"[ERROR] 가이드 '{var_col}' 로드 실패: {guide_path} -> {e}", file=sys.stderr); return None
    if var_col not in guide_df.columns: print(f"[ERROR] GUIDE '{var_col}' 없음: {guide_path}", file=sys.stderr); return None

    # Series 상태에서 unique() 호출 후 list로 변환
    target_series = guide_df[var_col].dropna().astype(str).str.strip()
    target_series = target_series[target_series != ''] # 빈 문자열 제거
    target_columns = target_series.unique().tolist() # 수정된 코드

    if not target_columns: print(f"[WARN] 가이드 '{os.path.basename(guide_path)}'의 '{var_col}' 열에 유효 변수명 없음.")
    print(f"[INFO] 가이드 '{os.path.basename(guide_path)}'에서 '{var_col}' 기준 {len(target_columns)}개 최종 대상 열 로드.")
    return target_columns

def build_guide_map(guide_dir):
    """업종별 가이드 매핑 (_industry_ 패턴 기준, 연도/최신 무관)"""
    guide_map = {} # 최종 맵: {industry: path}
    print(f"--- 가이드 스캔 ({guide_dir}) ---")
    guide_files_processed_count = 0 # 확인한 파일 수
    found_industries_count = 0 # 매핑에 성공한 업종 수 (고유 기준은 map 크기로 확인)

    for fname in os.listdir(guide_dir):
        # 1. 엑셀 파일이고 임시 파일이 아닌지 확인
        if fname.lower().endswith('.xlsx') and not fname.startswith('~$'):
            guide_files_processed_count += 1

            # 2. 새로운 함수를 사용하여 업종명 추출
            industry = extract_industry_from_filename(fname) # 이 함수는 이미 정의되어 있어야 함

            if industry: # 업종명이 성공적으로 추출되었으면
                # 3. guide_map에 추가 (이미 존재하면 경로 덮어쓰기 - 마지막 파일 기준)
                current_path = os.path.join(guide_dir, fname)
                guide_map[industry] = current_path
                found_industries_count += 1 # 매핑된 파일 카운트 (고유 업종 수는 아님)
                # print(f"  매핑됨: '{industry}' <- '{fname}'") # 디버깅 필요 시 주석 해제

            # else: # 업종명 추출 실패 시 무시 (또는 경고 로깅 가능)
                 # print(f"[INFO] 가이드 파일에서 업종명 추출 실패 또는 패턴 없음: {fname}")


    print(f"--- 가이드 스캔 완료 ---")
    print(f"총 {guide_files_processed_count}개 xlsx 파일 확인.")
    # 최종적으로 guide_map에 몇 개의 고유한 업종이 매핑되었는지 출력
    print(f"총 {len(guide_map)}개 고유 업종에 대한 가이드 매핑 완료.")

    if not guide_map:
        print(f"[WARN] 유효한 가이드 파일을 찾지 못했습니다 (업종명 추출 가능 파일 없음): {guide_dir}")
    else:
        # 최종 매핑된 정보 일부만 로깅 (너무 많으면 주석 처리)
        log_count = 0
        for industry, path_data in guide_map.items():
             print(f"[INFO] 매핑 확인: '{industry}' -> '{os.path.basename(path_data)}'")
             log_count += 1
             if log_count >= 5 and len(guide_map) > 10: # 너무 많으면 일부만 출력
                 print(f"  ... (총 {len(guide_map)}개 중 일부만 표시)")
                 break

    return guide_map

def merge_case_insensitive_lvs(df):
    """대소문자 다른 LV/xLV 컬럼 병합 및 이름 표준화 (성능 개선 버전)"""
    print(f"  - 대소문자 다른 LV/xLV 컬럼 병합/표준화 시작...")
    original_cols = list(df.columns)
    std_name_to_originals = defaultdict(list)
    non_lv_cols = []

    for col in original_cols:
        standard_name = standardize_lv_name(col) # 여기서 표준화 수행
        if standard_name:
            std_name_to_originals[standard_name].append(col)
        else:
            non_lv_cols.append(col)

    data_for_new_df = {}
    merged_group_count = 0

    for std_name in sorted(std_name_to_originals.keys()):
        original_cols_in_group = std_name_to_originals[std_name]
        if len(original_cols_in_group) > 1:
            ordered_group = [c for c in original_cols if c in original_cols_in_group]
            merged_series = df[ordered_group[0]].copy()
            for i in range(1, len(ordered_group)):
                merged_series = merged_series.combine_first(df[ordered_group[i]])
            data_for_new_df[std_name] = merged_series # 표준 이름으로 저장
            merged_group_count += 1
        elif len(original_cols_in_group) == 1:
            original_col = original_cols_in_group[0]
            data_for_new_df[std_name] = df[original_col] # 표준 이름으로 저장

    for col in non_lv_cols:
         data_for_new_df[col] = df[col] # 원본 이름으로 저장

    merged_df = pd.DataFrame(data_for_new_df, index=df.index)

    print(f"  - LV/xLV 컬럼 병합/표준화 완료 ({merged_group_count}개 그룹 병합). 생성 후 열 개수: {len(merged_df.columns)}")

    final_ordered_columns = []
    processed_for_reorder = set()
    # 원본 순서 기준으로 최종 순서 결정
    for col in original_cols:
        # 이 컬럼이 표준화되었다면 표준 이름을 사용하고, 아니면 원본 이름 사용
        std_name = standardize_lv_name(col)
        target_name = std_name if std_name else col
        # 최종 생성된 merged_df에 해당 이름이 있고 아직 처리 안됐으면 추가
        if target_name in merged_df.columns and target_name not in processed_for_reorder:
            final_ordered_columns.append(target_name)
            processed_for_reorder.add(target_name)

    missing_cols = [c for c in merged_df.columns if c not in processed_for_reorder]
    if missing_cols:
         print(f"[WARN] merge_case_insensitive_lvs: 순서 재구성 중 누락된 컬럼 발견, 뒤에 추가: {missing_cols}")
         final_ordered_columns.extend(missing_cols)

    return merged_df[final_ordered_columns].copy()

def reorder_columns_by_guide(df, guide_base_order, lv_mapping, codes_in_priority_order):
    """열 순서를 재배치합니다. (df 컬럼은 표준화된 상태여야 함)"""
    # 입력 df는 merge_case_insensitive_lvs 를 거쳐 LV/xLV 이름이 표준화된 상태임
    # lv_mapping의 키와 codes_in_priority_order는 원본 코드 이름임
    # 따라서 이 함수 내부의 표준화 로직(standardize_lv_name 호출) 및 매핑이 중요함
    print(f"  - 열 순서 재배치 시작 (기준: 가이드 '{VAR_LABEL_COL}' 순서, LV 우선순위 적용)...")
    current_cols = list(df.columns) # 표준화된 컬럼 목록

    guide_base_order_cleaned = [str(col).strip() for col in guide_base_order if col]
    valid_guide_base_order = [col for col in guide_base_order_cleaned if col in current_cols] # 가이드 기준 순서 (df에 있는것만)

    present_std_lvs_in_priority = []
    original_to_std_map = {} # 원본 코드 -> 표준 이름 매핑

    # 우선순위 리스트(원본 코드)를 기반으로, 현재 df에 존재하는 표준 LV 이름 목록 생성
    for original_code in codes_in_priority_order: # 원본 코드 리스트 순회
         std_name = standardize_lv_name(original_code) # 원본 코드를 표준화
         if std_name and std_name in current_cols: # 표준화 가능하고, 결과 df에 존재하면
             if std_name not in present_std_lvs_in_priority: # 표준 이름 중복 방지
                 present_std_lvs_in_priority.append(std_name)
             if original_code not in original_to_std_map: # 첫 발견된 원본 코드 기준 매핑
                 original_to_std_map[original_code] = std_name

    std_to_original_map = {v: k for k, v in original_to_std_map.items()} # 표준 이름 -> 원본 코드 역매핑

    final_order = []; processed_cols = set()
    var_to_lv_map = defaultdict(list); lvs_without_ref_vars = []

    # 표준 LV 이름 기준으로 관련 변수 매핑 생성
    for std_lv_code in present_std_lvs_in_priority: # 우선순위 순서대로 표준 LV 이름 순회
        original_code = std_to_original_map.get(std_lv_code) # 매핑된 원본 코드 찾기
        if not original_code: continue

        # lv_mapping (키: 원본코드, 값: 원본 변수레이블 리스트) 사용
        related_vars_raw = lv_mapping.get(original_code, [])
        related_vars_cleaned = [str(var).strip() for var in related_vars_raw if var]

        first_related_var_in_guide_order = None
        for guide_col in valid_guide_base_order: # 현재 df에 있는 가이드 순서 기준
            if guide_col in related_vars_cleaned:
                first_related_var_in_guide_order = guide_col; break

        if first_related_var_in_guide_order:
            # 기준 변수에 표준 LV 코드 매핑
            var_to_lv_map[first_related_var_in_guide_order].append(std_lv_code)
        else:
            lvs_without_ref_vars.append(std_lv_code)

    # 1. 가이드 베이스 순서 + LV 삽입
    for col in valid_guide_base_order:
        if col in var_to_lv_map:
            lvs_to_insert = [lv for lv in present_std_lvs_in_priority if lv in var_to_lv_map[col]]
            for lv_code in lvs_to_insert:
                if lv_code not in processed_cols: final_order.append(lv_code); processed_cols.add(lv_code)
        if col not in processed_cols: final_order.append(col); processed_cols.add(col)

    # 2. 관련 변수 없던 LV 추가
    if lvs_without_ref_vars:
        for lv_code in lvs_without_ref_vars:
            if lv_code not in processed_cols: final_order.append(lv_code); processed_cols.add(lv_code)

    # 3. 누락 컬럼 추가
    remaining_cols = [col for col in current_cols if col not in processed_cols]
    if remaining_cols:
         print(f"[WARN] 재배치 후 누락된 컬럼 발견(오류 가능성), 맨 뒤에 추가: {remaining_cols}")
         final_order.extend(remaining_cols)

    final_columns_existing = [col for col in final_order if col in df.columns]
    if len(final_columns_existing) != len(df.columns):
        print(f"[WARN] 최종 열 개수 불일치! DF: {len(df.columns)}, 최종: {len(final_columns_existing)}")
        missing_final = [col for col in df.columns if col not in final_columns_existing]
        if missing_final:
            print(f"[WARN] 최종 순서에서 누락된 열 재추가: {missing_final}")
            final_columns_existing.extend(missing_final)

    print(f"  - 열 순서 재배치 완료. 최종 열 개수: {len(final_columns_existing)}")
    return df[final_columns_existing]


# --- Main 함수 ---
def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1. 가이드 매핑 빌드
    guide_map = build_guide_map(GUIDE_DIR)
    if not guide_map: print("[ERROR] 처리할 가이드 파일 없음.", file=sys.stderr); return
    print(f"--- GUIDE 매핑 완료 ({len(guide_map)}개 업종) ---")

        # 2. 입력 파일 스캔 및 그룹화
    industry_year_files = defaultdict(lambda: defaultdict(list))
    print(f"--- 입력 파일 스캔 ({NCSI_DATA_DIR}) ---")
    found_files_count = 0; skipped_files_count = 0
    for root, _, files in os.walk(NCSI_DATA_DIR):
        for fname in files:
            if fname.startswith('~$') or not fname[0].isdigit():
                skipped_files_count += 1; continue

            file_year = fname[:2]
            if file_year in TARGET_YEARS and fname.lower().endswith(('.csv', '.xlsx', '.xls')):

                # 새로운 함수를 사용하여 입력 파일의 업종명 추출
                potential_industry = extract_industry_from_filename(fname)

                # 추출된 업종명이 guide_map의 키(가이드에서 추출된 업종명)에 있는지 확인
                if potential_industry is not None and potential_industry in guide_map:
                    matched_industry = potential_industry # 매칭 성공
                    full_path = os.path.join(root, fname)
                    industry_year_files[matched_industry][file_year].append(full_path)
                    found_files_count += 1
                else:
                    # 업종명 추출 실패 또는 guide_map에 없는 경우 건너뜀
                    skipped_files_count +=1

            else:
                # 연도 또는 확장자가 맞지 않는 경우 건너뜀
                skipped_files_count += 1

    if found_files_count == 0 : print(f"[ERROR] 처리 대상 파일 없음 (업종 매칭 실패 포함).", file=sys.stderr); return
    else: print(f"--- {found_files_count}개 파일 분류 완료 (건너뜀: {skipped_files_count}개) ---")


    print(f"--- 업종별 데이터 처리 시작 ({len(industry_year_files)}개 업종) ---")
    # 3. 업종별 데이터 처리
    processed_industries_count = 0
    error_industries = []

    for industry, year_files_map in industry_year_files.items():
        print(f"\n=== 업종 처리 시작: {industry} ===")
        guide_path = guide_map.get(industry)
        if not guide_path:
            print(f"[ERROR] 가이드 경로 없음: {industry}. 건너뜁니다.", file=sys.stderr)
            error_industries.append(industry + " (가이드 없음)")
            continue

        # --- [수정됨] 사전 검증 단계 (부분 표준화된 이름 기준) ---
        print("  - 사전 검증 시작: 가이드/데이터 컬럼 부분 표준화 후 확인...")
        # 가이드에서 LV/xLV만 표준화된 레이블 목록 가져오기
        required_val_labels, forbidden_val_labels = get_validation_labels_from_guide(
            guide_path, LV_LABEL_GUIDE_COLUMNS, X_LV_LABEL_GUIDE_COLUMNS
        )
        if required_val_labels is None or forbidden_val_labels is None:
            print(f"[ERROR] 가이드 검증 레이블 로드 실패: {industry}. 건너뜁니다.", file=sys.stderr)
            error_industries.append(industry + " (가이드 검증 레이블 로드 실패)")
            continue

        # 해당 업종의 모든 대상 연도 파일 경로 수집
        all_input_files_for_industry = []
        for year in TARGET_YEARS:
            if year in year_files_map:
                all_input_files_for_industry.extend(year_files_map[year])
        if not all_input_files_for_industry:
             print(f"[WARN] 업종 '{industry}'에 대한 입력 파일 없음. 건너뜁니다.")
             continue

        # 모든 입력 파일에서 LV/xLV만 표준화된 컬럼명 집합 가져오기
        actual_val_columns = get_combined_columns_from_files(all_input_files_for_industry)
        if not actual_val_columns:
             print(f"[ERROR] 업종 '{industry}'의 입력 파일에서 컬럼 정보를 읽을 수 없음. 건너뜁니다.", file=sys.stderr)
             error_industries.append(industry + " (컬럼 읽기 실패)")
             continue

        # LV 검증 (부분 표준화된 이름 기준)
        missing_lv_labels = required_val_labels - actual_val_columns
        if missing_lv_labels:
            print(f"[ERROR] LV 변수 누락 에러 ({industry}): 다음 필수 레이블(또는 표준화된 LV)이 입력 파일에 없습니다: {sorted(list(missing_lv_labels))}", file=sys.stderr)
            error_industries.append(industry + " (모델값 확인안됨)")
            continue

        # xLV 검증 (부분 표준화된 이름 기준)
        existing_xlv_labels = forbidden_val_labels.intersection(actual_val_columns)
        if existing_xlv_labels:
            print(f"[ERROR] xLV 변수 사전 존재 에러 ({industry}): 다음 레이블(또는 표준화된 xLV)은 입력 파일에 미리 존재하면 안됩니다: {sorted(list(existing_xlv_labels))}", file=sys.stderr)
            error_industries.append(industry + " (비모델값 확인됨)")
            continue

        print("  - 사전 검증 통과.")
        # --- 사전 검증 단계 끝 ---


        # --- 검증 통과 후, 기존 로직 수행 ---
        # 가이드 정보 로드 (LV/xLV 처리 및 최종 열 선택용)
        lv_processing_mapping, codes_in_guide_priority_order = load_lv_mapping(
            guide_path, VAR_LABEL_COL, LV_LABEL_GUIDE_COLUMNS + X_LV_LABEL_GUIDE_COLUMNS
        )
        target_base_columns = load_target_columns_from_guide(guide_path, VAR_LABEL_COL)

        if lv_processing_mapping is None or target_base_columns is None:
            print(f"[ERROR] 처리용 가이드 정보 로드 실패: {industry}. 건너<0xEB><0x9B><0x84>니다.", file=sys.stderr)
            error_industries.append(industry + " (처리 가이드 로드 실패)")
            continue

        # --- 데이터 로드 및 병합 (Concat) ---
        dfs_to_merge = []
        sorted_years = sorted([y for y in year_files_map.keys() if y in TARGET_YEARS])
        print(f"  - 데이터 로딩 시작 (연도순: {sorted_years})")
        load_success = True
        for year in sorted_years:
            if year in year_files_map:
                for file_path in year_files_map[year]:
                    df_loaded = load_data_file(file_path) # 컬럼명 strip됨
                    if df_loaded is not None and isinstance(df_loaded, pd.DataFrame):
                         if df_loaded.columns.has_duplicates:
                             duplicates = df_loaded.columns[df_loaded.columns.duplicated()].unique()
                             print(f"[WARN] 파일 '{os.path.basename(file_path)}'에 중복된 컬럼명이 있습니다: {list(duplicates)}. 병합 시 문제가 발생할 수 있습니다.")
                         df_loaded['_SOURCE_YEAR_'] = int(year)
                         dfs_to_merge.append(df_loaded)
                    elif df_loaded is None:
                         print(f"[ERROR] 데이터 병합 중 파일 로드 실패: {os.path.basename(file_path)}. 이 업종 처리를 중단합니다.", file=sys.stderr)
                         load_success = False
                         break
            if not load_success: break
        if not load_success:
             error_industries.append(industry + " (데이터 로드 실패)")
             continue

        if not dfs_to_merge:
            print(f"[WARN] 병합할 데이터 파일 없음: {industry}. 건너<0xEB><0x9B><0x84>니다.")
            continue
        print(f"  - 데이터 병합(Concat) 시작... ({len(dfs_to_merge)}개)")
        try:
            # 병합된 DataFrame 컬럼은 아직 원본(strip된) 상태
            merged_df_raw = pd.concat(dfs_to_merge, ignore_index=True, sort=False, join='outer')
            merged_df_raw.columns = [str(col).strip() for col in merged_df_raw.columns] # 최종 확인
            print(f"  - 병합 후 크기: {merged_df_raw.shape}")
        except Exception as e:
            print(f"[ERROR] 데이터 병합(Concat) 실패: {industry} -> {e}", file=sys.stderr)
            error_industries.append(industry + " (Concat 실패)")
            continue

        # --- 행 정렬 ---
        print("  - 행 정렬 시작..."); merged_df_raw = merged_df_raw.sort_values(by='_SOURCE_YEAR_', ascending=True, kind='stable'); print("  - 행 정렬 완료.")
        merged_df_raw = merged_df_raw.drop(columns=['_SOURCE_YEAR_'])

        # --- [수정됨] xLV 생성 (표준 이름으로) ---
        print("  - xLV 생성 시작 (가이드 매핑 기준)..."); generated_xlv_count = 0
        merged_df_with_xlv = merged_df_raw.copy() # 원본 유지 위해 복사
        if lv_processing_mapping:
            for code, vars_list in lv_processing_mapping.items(): # code는 원본 LV/xLV 이름
                standard_code_name = standardize_lv_name(code) # 생성할 표준 이름
                if standard_code_name and standard_code_name.startswith('xLV') and standard_code_name not in merged_df_with_xlv.columns:
                    # vars_list는 원본 변수 레이블 (strip됨)
                    valid_cols = [v for v in vars_list if v in merged_df_with_xlv.columns] # 원본 레이블로 컬럼 찾기
                    if valid_cols:
                        try:
                            numeric_df = merged_df_with_xlv[valid_cols].apply(pd.to_numeric, errors='coerce')
                            # 표준 이름으로 새 컬럼 추가
                            merged_df_with_xlv[standard_code_name] = numeric_df.mean(axis=1, skipna=True)
                            generated_xlv_count += 1
                        except Exception as e:
                            print(f"    - [WARN] xLV 생성 실패 ({standard_code_name} <- {valid_cols}): {e}")
        print(f"  - xLV 생성 완료 ({generated_xlv_count}개 생성). 생성 후 크기: {merged_df_with_xlv.shape}")

        # --- [수정됨] 대소문자 다른 LV/xLV 컬럼 병합/표준화 ---
        # 입력: xLV가 추가되었고, 여전히 원본 LV 컬럼(lv1, LV1 등)이 있는 DataFrame
        # 출력: 모든 LV/xLV 컬럼이 표준 이름으로 변경/병합된 DataFrame
        merged_lv_std_df = merge_case_insensitive_lvs(merged_df_with_xlv)

        # --- 최종 열 선택 (가이드 기준 변수 + 표준화된 LV/xLV) ---
        print(f"  - 최종 포함할 열 필터링 시작...")
        # 최종 포함 대상 컬럼 집합 생성
        # 1. 가이드 '변수레이블' (원본, strip됨)
        final_columns_allowed_set = set(target_base_columns)
        # 2. 가이드에서 정의된 모든 LV/xLV의 표준 이름 추가 (lv_processing_mapping 사용)
        for code in lv_processing_mapping.keys():
            std_name = standardize_lv_name(code)
            if std_name: final_columns_allowed_set.add(std_name)

        print(f"    - 최종 허용 컬럼 개수 (이론상): {len(final_columns_allowed_set)}")

        # merged_lv_std_df 에서 허용된 컬럼만 선택 (이제 모든 컬럼 이름이 목표 형태)
        filtered_cols = [col for col in merged_lv_std_df.columns if col in final_columns_allowed_set]
        filtered_df = merged_lv_std_df[filtered_cols]
        print(f"  - 최종 포함 열 필터링 완료. 필터링 후 열 개수: {len(filtered_df.columns)}")

        # --- 최종 열 순서 재배치 ---
        # 입력: 필터링 + 표준화 완료된 DataFrame (filtered_df)
        # 기준: target_base_columns (원본 변수레이블), lv_processing_mapping (원본 코드->원본 변수), codes_in_guide_priority_order (원본 코드)
        final_output_df = reorder_columns_by_guide(filtered_df, target_base_columns, lv_processing_mapping, codes_in_guide_priority_order)

        # --- 결과 저장 ---
        output_filename = f"KPC_NCSI_{industry}_DATA.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        try:
            final_output_df.to_excel(output_path, index=False)
            print(f"[SUCCESS] 최종 결과 파일 저장 완료: {output_path}")
            print(f"  - 최종 데이터 크기 (행, 열): {final_output_df.shape}")
            processed_industries_count += 1
        except Exception as e:
            print(f"[ERROR] 최종 파일 저장 실패: {output_path} -> {e}", file=sys.stderr)
            error_industries.append(industry + " (저장 실패)")

        print(f"=== 업종 처리 완료: {industry} ===")

    print(f"\n--- 모든 작업 완료 ---")
    print(f"총 {len(industry_year_files)}개 업종 시도.")
    print(f"  - 성공적으로 처리 및 저장된 업종 수: {processed_industries_count}")
    if error_industries:
        print(f"  - 오류 또는 건너뛴 업종 수: {len(error_industries)}")
        print(f"  - 오류/건너뛴 업종 목록: {error_industries}")
    else:
        print(f"  - 모든 업종 처리 성공.")

if __name__ == '__main__':
    print("--- Configuration ---")
    print(f"Input Directory: {NCSI_DATA_DIR}"); print(f"Guide Directory: {GUIDE_DIR}"); print(f"Output Directory: {OUTPUT_DIR}")
    print(f"Variable Label Column: {VAR_LABEL_COL}");
    print(f"LV Label Guide Columns (Indices, Required): {LV_LABEL_GUIDE_COLUMNS}");
    print(f"xLV Label Guide Columns (Indices, Forbidden): {X_LV_LABEL_GUIDE_COLUMNS}");
    print(f"Target Years: {TARGET_YEARS}")
    print("---------------------")
    main()