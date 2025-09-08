import os
import pandas as pd

# === 설정 ===
target_folder = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Result"    #수정 필요



required_columns = ['id', 'year', 'sector', 'industry', 'firm', 'firm1', 'area', 'gender', 'age', 'age1']

# === 검사 함수 ===
def check_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = [str(col).strip().lower() for col in df.columns]

        missing_columns = []
        empty_columns = []

        for col in required_columns:
            if col not in df.columns:
                missing_columns.append(col)
            elif df[col].dropna().empty:
                empty_columns.append(col)

        return missing_columns, empty_columns
    except Exception as e:
        print(f"❌ 파일 읽기 실패: {os.path.basename(file_path)} → {e}")
        return None, None

# === 메인 ===
def main():
    print("QA 검사 시작...\n")
    files_checked = 0
    errors_found = []

    for fname in os.listdir(target_folder):
        if fname.lower().endswith(".xlsx") and not fname.startswith("~$"):
            file_path = os.path.join(target_folder, fname)
            missing, empty = check_excel_file(file_path)
            files_checked += 1

            if missing or empty:
                errors_found.append({
                    '파일명': fname,
                    '누락된 열': missing,
                    '값이 없는 열': empty
                })

    print(f"\n📁 총 검사 파일 수: {files_checked}")
    if errors_found:
        print(f"⚠ 오류 발생 파일 수: {len(errors_found)}\n")
        for err in errors_found:
            print(f"📌 파일: {err['파일명']}")
            if err['누락된 열']:
                print(f"  - 누락된 열: {', '.join(err['누락된 열'])}")
            if err['값이 없는 열']:
                print(f"  - 값이 없는 열: {', '.join(err['값이 없는 열'])}")
            print()
    else:
        print("✅ 모든 파일이 정상입니다.")

if __name__ == "__main__":
    main()
