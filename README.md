# 🦄 GIST 새내기배움터 조 자동 배정 프로그램 🚀

<div align="center">
  <img src="https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/Rich-UI-009900?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Pandas-Data-150458?style=for-the-badge&logo=pandas&logoColor=white" />
</div>

<br>

**지스트 문화행사위원회**에서 사용하는 새내기배움터 조 편성 자동화 프로그램입니다.  
수백 명의 학생들을 **학과(학부), 성비, 신캠 조, 나이** 등을 종합적으로 고려하여 최적의 조 배정을 진행합니다.

## ✨ 배정 로직

### **🧬 1. 학과(학부)**
도전탐색과정 학생과 반도체공학과 학생들이 있을 것인데, 이 학생들을 조별로 고르게 배정합니다.

### **🚻 2. 성비**
조마다 성비가 서로 비슷하게 되도록 배정합니다.

### **🏕️ 3. 신캠 조**
특정 새터 조에 특정 신캠 조 학생들이 쏠리는 것을 방지하여 배정합니다.
단, 아는 사람이 너무 없는 것을 방지하기 위하여 최대한 적당한 인원 (2~3명)씩 묶어서 배정합니다.
(그럼에도 1명씩 배정되는 새내기가 있을 수 있습니다.)

### **🔢 4. 나이**
특정 조에 나이가 적거나 많은 사람이 쏠리는 것을 방지하여 배정합니다.
또한, 나이가 많은 새내기의 경우, 최소 한 명 이상의 조장이 해당 새내기보다 나이가 많거나 같게끔 배정합니다.
(단, 해당 학생이 새터 전체 조장들보다 나이가 많은 경우, 나이가 가장 많은 조장들의 새터 조 중에 한 조에 배정됩니다.)

## 🛠️ 설치 및 실행 방법

이 프로그램은 개인정보 보호를 위해 **데이터 파일(xlsx)을 포함하고 있지 않습니다.** 관리자가 소유한 `연도ST_leader.xlsx` 및 `연도ST_freshmen.xlsx` 파일이 필요합니다.

1. **저장소 클론**
   ```bash
   git clone [https://github.com/YOUR-USERNAME/GIST-Saeteo-TeamBuilder.git](https://github.com/YOUR-USERNAME/GIST-Saeteo-TeamBuilder.git)
   cd GIST-Saeteo-TeamBuilder
   ```

2. 라이브러리 설치
    ```bash
    pip install -r requirements.txt
    ```
*(프로그램 실행 시 자동으로 설치를 시도합니다)*

3. 실행
    ```bash
    python main.py
    ```

## ⚠️ 주의사항
* 본 프로그램은 로컬 환경에서 엑셀 데이터를 읽고 처리하며, 외부로 데이터를 전송하지 않습니다.
* 생성된 결과 파일(team_result_...xlsx)에는 개인정보가 포함될 수 있으므로 관리에 유의하십시오.

---
## 🛠️ 히스토리
* v26.1.21: Release!
---
Developed by **GIST 문화행사위원회**