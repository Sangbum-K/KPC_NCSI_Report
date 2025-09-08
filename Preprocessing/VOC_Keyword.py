import os
import pandas as pd
import re
from konlpy.tag import Okt

# 1. 설정
target_path = r"C:\Users\dwjung\Desktop\Workspace\Program\Target\Result"  # 수정 필요

okt = Okt()
custom_stopwords = ['너무', '정말', '진짜', '자다', '약간','조금','별로', '좀']

# 2. 키워드 추출 함수
def extract_keywords(text):
    if pd.isna(text):
        return ""

    # '없다', '되다' 관련 표현 전처리
    processed_text = re.sub(r'없(어요|다|습니다|으면|어서|음)?', '없다', str(text))
    processed_text = re.sub(r'되었(다|어요|습니다)', '되다', processed_text)

    # 1. 원형 복원 (stem=True)
    morphs = okt.pos(processed_text, stem=True)

    # 2. 불용 품사 및 사용자 정의 불용어를 먼저 제거
    stop_tags = ['Josa', 'Eomi', 'Punctuation', 'Suffix', 'Conjunction', 'PreEomi', 'Foreign']
    filtered = [m for m in morphs if m[1] not in stop_tags and m[0] not in custom_stopwords]

    keywords = []
    noun_buffer = []
    i = 0
    while i < len(filtered):
        word, tag = filtered[i]

        # 3. '안' + Verb 패턴을 먼저 감지 (Okt가 '안'을 Noun으로 오분류하는 경우 대비)
        if word == '안' and (i + 1 < len(filtered) and filtered[i+1][1] == 'Verb'):
            if noun_buffer:
                keywords.append("".join(noun_buffer))
                noun_buffer = []
            
            # '안'과 뒤따르는 동사를 합쳐서 키워드로 만듦
            next_word, _ = filtered[i+1]
            keywords.append('안' + next_word)
            i += 2
            continue

        # 4. 일반 용언 + '않다' 부정 패턴 감지 (예: 많지 않다, 다양하지 않다)
        if tag in ['Adjective', 'Verb'] and word.endswith('다'):
            if (i + 1 < len(filtered) and filtered[i+1][0] == '않다' and filtered[i+1][1] == 'Verb'):
                
                # 명사 버퍼가 있으면 먼저 처리
                if noun_buffer:
                    keywords.append("".join(noun_buffer))
                    noun_buffer = []
                
                # '...지않다' 형태로 재구성 (예: 많다 -> 많지않다)
                base_form = word[:-1]
                reconstructed_word = base_form + '지않다'
                keywords.append(reconstructed_word)
                i += 2  # 2개 토큰을 처리했으므로 인덱스 점프
                continue

        # 5. (패턴 아닐 시) 명사는 버퍼에 추가
        if tag == 'Noun':
            noun_buffer.append(word)
        # 6. 그 외 용언 및 기타 토큰 처리
        else:
            if noun_buffer:
                keywords.append("".join(noun_buffer))
                noun_buffer = []
            
            if tag in ['Adjective', 'Verb']:
                 keywords.append(word)
        
        i += 1
    
    # 6. 마지막에 남은 명사 버퍼 처리
    if noun_buffer:
        keywords.append("".join(noun_buffer))

    # 7. 최종 중복 제거 (순서 유지)
    return " ".join(list(dict.fromkeys(keywords)))


# 3. 의미 없는 응답 제거
meaningless_set = {'없음', '모름', '무응답', '없다'}

def is_meaningless_only(text):
    if pd.isna(text):
        return False
    tokens = re.split(r'[/, ]+', str(text).strip())
    return all(tok in meaningless_set for tok in tokens if tok)

# 5. 엑셀 파일 반복 처리

for filename in os.listdir(target_path):
    if filename.endswith(".xlsx"):
        path = os.path.join(target_path, filename)
        try:
            df = pd.read_excel(path)

            # 열 중 하나라도 전체가 공백/NaN이면 스킵
            empty_columns = [
                col for col in df.columns[:9]
                if df[col].dropna().astype(str).str.strip().eq("").all()
            ]

            
            """

            if empty_columns:
                print(f"❌ 데이터 누락: {filename}")
                for col in empty_columns:
                    print(f"   - 공백 열: '{col}'")
                continue
            """
      

            target_col = df.columns[8]

            # 키워드 추출
            df['키워드 문장'] = df[target_col].apply(extract_keywords)

            # 의미 없는 응답 제거
            df['키워드 문장'] = df.apply(
                lambda row: "" if is_meaningless_only(row[target_col]) else row['키워드 문장'],
                axis=1
            )

            df.to_excel(path, index=False)
            print(f"✅ 성공: {filename}")

        except Exception as e:
            print(f"❌ 오류 발생: {filename} → {e}")


# 6. 테스트
#print(extract_keywords(('상담사가 좀 불친절해요')))

