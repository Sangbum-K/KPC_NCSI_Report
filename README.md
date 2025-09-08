# KPC_NCSI_Report
**CX Data Visualization Project**

---

## 프로젝트 목표 
- 데이터 모델링
  - 고객 경험 관점에서 데이터 수집 및 관리 체계 점검
  - 기존 데이터 표준화, 재처리 등 필요 데이터 생성
  - 수집 데이터 활용(관리포인트, 인사이트) 방안 수립
    
- 데이터 시각화
  - 지속적 관리 가능한 시각화 툴 제공
 
<br>

---

## 데이터 전처리
| 스크립트      | 기능        | 설명         | 
|--------------|-------------|----------------|
| SAV_to_Xlsx.py | SPSS->SAV 변환 | SPSS 파일(.SAV)을 Excel 형식(.xlsx)으로 변환 |
| Merge.py | 데이터 병합/정렬 | 모델값, 비모델값 분류, 다년도 데이터 통합, 자동 컬럼 매핑 |
| VOC_Extract | VOC 추출 | SAV 파일에서 VOC(불만족 의견) 텍스트 추출하여 Excel 저장 |
| VOC_Keyword | VOC 키워드화 | 형태소 분석을 통한 불용어 제거 및 키워드 추출 |

<br>

---

##  데이터 구조
```
KPC_NCSI_Report/Data
│
├── 기준데이터/
│ ├── KPC_NCSI_국가단위Data
│ ├── KPC_NCSI_경제단위Data
│ ├── KPC_NCSI_업종단위Data
│ ├── KPC_NCSI_기업단위Data
│ └── KPC_NCSI_업종&기업 코드
│
├── 품졸요인데이터/
│ └── 변수가이드_(기업명)(업종코드)    # 설문 항목 분류와 값 유형 구분을 위한 데이터 스키마 가이드
│
├── 조사데이터/
│ └── KPC_NCIS(기업명)Data           # 실 데이터
│
├── VOC 데이터/
  └── KPC_NCSI_VOC(기업명)
```

---

## ERD

![ERD](https://github.com/Sangbum-K/KPC_NCSI_Report/blob/main/EDR.PNG)

---

## 시현 영상

![Demo](https://github.com/Sangbum-K/KPC_NCSI_Report/blob/main/Report%20-%20Demo.gif)

