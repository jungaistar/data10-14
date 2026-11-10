// 로그인 설정 — 이 파일 한 곳만 바꾸면 된다(생성: data10-00/scripts/apply-auth-kit.mjs, 대문 문구 출처 docs/plan.json).
// 내 Supabase 로 옮길 때: url · anonKey 를 내 프로젝트 값으로, 카카오를 설정하지 않았으면 kakao: false. 안내 = docs/내_깃허브로_옮기기.md
// prefix 는 supabase/auth.sql 의 표 이름 앞부분과 같아야 한다.
window.AUTH_CONFIG = {
  "url": "https://hcmgdztsgjvzcyxyayaj.supabase.co",
  "anonKey": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhjbWdkenRzZ2p2emN5eHlheWFqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE0MzU4ODcsImV4cCI6MjA4NzAxMTg4N30.gznaPzY1l8qDAPsEyYNR9KS7f7VqS3xaw-_2HTSwSZw",
  "prefix": "data1014",
  "kakao": true,
  "idleMinutes": 30,
  "localBypass": true,
  "site": {
    "title": "재고실사 Start-End",
    "summary": "ERP 재고 엑셀과 창고 배정표로 기준 일자·시각이 박힌 재고실사 기록서를 만들고, 휴대폰이나 SharePoint 엑셀로 적은 실물수량·사유를 합쳐 차이대사·분석과 버전이 있는 결과 보고서(PPT·엑셀)까지 잇는 브라우저 도구입니다.",
    "course": "현장 데이터 디지털화 전문가과정 · 2차수",
    "features": [
      {
        "name": "작업지시 · 기록서",
        "desc": "ERP 재고 엑셀과 창고 배정표로 기준 일자·시각이 박힌 작업지시와 SharePoint 취합용 기록서 엑셀을 만듭니다."
      },
      {
        "name": "휴대폰 현장 입력",
        "desc": "실물수량을 적고 차이가 있으면 사유를 꼭 고릅니다. 사진·특이사항을 남기고 확인 시각·작성자는 자동입니다."
      },
      {
        "name": "취합 · 차이대사",
        "desc": "SharePoint 파일과 휴대폰 결과를 합치고 충돌을 고른 뒤, 수량·금액 차이와 사유 미입력을 봅니다."
      },
      {
        "name": "보고서 PPT · 결과서",
        "desc": "직전 보고서 pptx 에 내용만 바꿔 넣고(상세 표가 여러 장이어도 장마다 표가 채워짐), 버전·기록 ID 가 있는 결과서 엑셀·Word 를 냅니다."
      },
      {
        "name": "분석(BI)",
        "desc": "창고별·창고위치(사내/사외)별·실사자별·사유별·시간대별 진행률·정확도·차이 금액. 표에서 고른 행만 그래프로 그립니다."
      },
      {
        "name": "진행 현황",
        "desc": "진행률·정확도·차이 품목·순차이, 실사자별 진행(배정 기준과 직접 센 품목 기준), 제외 창고, 미해결 충돌을 보여 줍니다."
      }
    ],
    "steps": [
      "로그인 · 승인",
      "작업지시",
      "기록서 배포",
      "현장 입력",
      "취합 · 충돌",
      "차이대사 · 분석"
    ],
    "flow": [
      {
        "step": "로그인 · 승인",
        "desc": "가입 · 사용 신청 뒤 관리자가 승인하면 도구가 열림"
      },
      {
        "step": "작업지시",
        "desc": "ERP 재고 엑셀 · 기준 일자·시각 · 창고 배정표로 실사 회차를 만듦"
      },
      {
        "step": "기록서 배포",
        "desc": "기록서 엑셀을 SharePoint 에 올리거나 휴대폰에 넣음, 인쇄본도 가능"
      },
      {
        "step": "현장 입력",
        "desc": "실물수량 · 사유(차이가 있으면 필수) · 특이사항 · 사진"
      },
      {
        "step": "취합 · 충돌",
        "desc": "SharePoint 파일 · 휴대폰 결과를 합치고 다른 수량은 골라 채택"
      },
      {
        "step": "차이대사 · 분석",
        "desc": "수량·금액 차이, 사유, 창고·창고위치·실사자·사유별 분석"
      }
    ],
    "goals": [
      "ERP 재고 엑셀을 올리는 것에서 작업지시를 출발하고, 기준 일자·시각을 반드시 남깁니다.",
      "현장에서 휴대폰이나 SharePoint 엑셀로 실물수량·사유·특이사항·사진을 남기고, 확인 시각·작성자는 자동으로 붙입니다.",
      "여러 사람의 결과를 한곳에 합쳐 차이대사와 창고별·창고위치별·실사자별·사유별 분석을 바로 봅니다."
    ],
    "screen": "docs/img/screen.png",
    "plan": "plan.html",
    "home": true,
    "landing": {
      "image": "images/landing.jpg",
      "imageMode": "cover",
      "colors": {
        "navy": "#005EB8",
        "navy2": "#003b75",
        "brand": "#005EB8",
        "brandDark": "#004a91",
        "accent": "#a9d0f5"
      },
      "company": {
        "name": "두산",
        "en": "DOOSAN"
      }
    }
  }
};
