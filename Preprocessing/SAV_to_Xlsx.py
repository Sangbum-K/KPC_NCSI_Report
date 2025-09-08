import os
import pyreadstat
import pandas as pd

# === 설정 ===
target_folder = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Data"          #수정 필요    
result_folder = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Result"        #수정 필요
os.makedirs(result_folder, exist_ok=True)


# === 변환 함수 ===
def convert_sav_to_excel(file_path, result_path):
    try:
        df, meta = pyreadstat.read_sav(file_path)
        df.to_excel(result_path, index=False)
        print(f"✅ 변환 성공: {os.path.basename(file_path)}")
        return True
    except Exception as e:
        print(f"❌ 변환 실패: {os.path.basename(file_path)} → {e}")
        return False

# === 전체 실행 ===
def main():
    print("SAV → Excel 일괄 변환 시작...\n")

    success_count = 0
    fail_count = 0

    for fname in os.listdir(target_folder):
        if fname.lower().endswith(".sav"):
            file_path = os.path.join(target_folder, fname)
            result_fname = os.path.splitext(fname)[0] + ".xlsx"
            result_path = os.path.join(result_folder, result_fname)

            if convert_sav_to_excel(file_path, result_path):
                success_count += 1
            else:
                fail_count += 1

    print("\n📌 변환 요약")
    print(f" - 성공: {success_count}개")
    print(f" - 실패: {fail_count}개")
    print("처리 완료.")

if __name__ == "__main__":
    main()
