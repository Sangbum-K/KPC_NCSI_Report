import os
import pyreadstat
import pandas as pd
from collections import defaultdict

# 입력 및 출력 폴더 설정
input_folder = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Data"
output_folder = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Result"

os.makedirs(output_folder, exist_ok=True)

# 파일을 기업명 기준으로 그룹핑
company_files = defaultdict(list)
for fname in os.listdir(input_folder):
    if fname.lower().endswith(".sav"):
        parts = fname.split("_")
        if len(parts) >= 3:
            company = parts[2]
            company_files[company].append(fname)

# 병합 결과 통계
success_count = 0
fail_count = 0

# 기업별로 파일 병합
for company, files in company_files.items():
    print(f"\n병합: {company} ({len(files)}개 파일)")
    print(f"    포함 파일:")
    for f in files:
        print(f"        - {f}")

    merged_rows = []
    industry_code = None  # 업종 코드 추출용

    for fname in files:
        full_path = os.path.join(input_folder, fname)

        encodings_to_try = [None, 'cp949', 'euc-kr', 'utf-8', 'latin1']
        df, meta = None, None


        
        for encoding in encodings_to_try:
            try:
                if encoding is None:
                    df, meta = pyreadstat.read_sav(full_path)  # 기본 인코딩
                else:
                    df, meta = pyreadstat.read_sav(full_path, encoding=encoding)
                print(f"    ✅ 성공 (인코딩: {encoding or '기본'})")
                break
            except Exception as e:
                if encoding == encodings_to_try[-1]:  # 마지막 시도
                    print(f"    ❌ {fname} 읽기 오류: {e}")
                    continue
                else:
                    continue

        if df is None:
            continue        

        # 소문자 → 원본 컬럼명 매핑
        col_map = {col.lower(): col for col in df.columns}

        # 필수 열 확인
        if 'id' not in col_map or 'industry' not in col_map or 'firm' not in col_map:
            print(f"    ⚠ {fname}: 필수 열(id, industry, firm) 중 일부 누락. 스킵.")
            continue

        # 업종 코드 추출
        industry_col = col_map['industry']
        if industry_code is None:
            industry_values = df[industry_col].dropna().unique()
            if len(industry_values) == 0:
                print(f"    ⚠ {fname}: 업종코드 값 없음. 스킵.")
                continue
            industry_code = industry_values[0]

        # dissat 열 추출
        dissat_cols = [col for col in df.columns if col.lower().startswith('dissat')]
        if not dissat_cols:
            print(f"    ⚠ {fname}: dissat 관련 열 없음. 스킵.")
            continue

        # 타겟 열 정의
        target_cols = ['id', 'year', 'firm', 'firm1', 'area', 'gender', 'age', 'age1']
        selected_cols = {}
        for col in target_cols:
            if col not in col_map:
                print(f"    ⚠ {fname}: 열 '{col}' 없음. 해당 열은 None으로 대체.")
                selected_cols[col] = None
            else:
                selected_cols[col] = col_map[col]

        # 각 dissat 열 처리
        for dissat_col in dissat_cols:
            labelset_name = meta.variable_value_labels.get(dissat_col, None)
            if isinstance(labelset_name, str):
                matched_label_dict = meta.value_labels.get(labelset_name, {})
            elif isinstance(labelset_name, dict):
                matched_label_dict = labelset_name
            else:
                matched_label_dict = {}

            sub_df = pd.DataFrame()
            for col in target_cols:
                real_col = selected_cols[col]
                sub_df[col] = df[real_col] if real_col else None
            sub_df['dissat'] = df[dissat_col].map(matched_label_dict)

            # 열 순서 고정
            sub_df = sub_df[target_cols + ['dissat']]
            merged_rows.append(sub_df)

    # 결과 저장
    if merged_rows:
        final_df = pd.concat(merged_rows, ignore_index=True)
        industry_code_str = str(int(float(industry_code)))
        output_filename = f"{company}_{industry_code_str}_VOC.xlsx"
        output_path = os.path.join(output_folder, output_filename)
        try:
            final_df.to_excel(output_path, index=False)
            print(f"  ✅ 저장 완료: {output_filename}")
            success_count += 1
        except Exception as e:
            print(f"  ❌ 저장 실패: {e}")
            fail_count += 1
    else:
        print(f"  ❌ 병합 실패: 유효한 데이터 없음")
        fail_count += 1

# 요약 출력
print("\n📊 병합 요약")
print(f"  ✅ 성공: {success_count}개")
print(f"  ❌ 실패: {fail_count}개")
