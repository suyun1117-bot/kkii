# JSON 데이터 직렬화 및 역직렬화(API 응답 파싱 등)를 위한 표준 라이브러리
import json
# 벡터 크기 계산(L2 norm) 및 제곱근 연산을 위한 수학 라이브러리
import math
# 운영체제 환경변수 접근 및 파일 경로 검사를 위한 라이브러리
import os
# 정규표현식을 통한 문자열 정제 및 패턴 매칭 라이브러리
import re
# 캐시 만료 시간(TTL) 계산 및 타임스탬프 측정을 위한 시간 라이브러리
import time
# 동적으로 파이썬 모듈을 가져오기 위한 라이브러리
import importlib
# 특정 모듈의 설치 여부(spec)를 사전에 확인하기 위한 유틸리티
import importlib.util
# 단어 및 토큰 빈도수 계산(BoW 벡터 생성)을 위한 Counter 클래스
from collections import Counter

# 엑셀 규정 데이터(.xlsx) 처리 및 데이터프레임 조작 라이브러리
import pandas as pd
# LLM 서버(Groq/OpenAI)로 HTTP 스트리밍 요청을 전송하기 위한 라이브러리
import requests
# 웹 인터페이스(UI) 및 세션 상태 관리를 위한 Streamlit 프레임워크
import streamlit as st

# LangChain의 프롬프트 템플릿 생성 모듈
from langchain_core.prompts import ChatPromptTemplate
try:
    # 외부 웹 검색을 위한 Tavily 클라이언트 로드 시도
    from tavily import TavilyClient
except Exception:
    # Tavily 라이브러리가 미설치되었을 경우 None으로 대체
    TavilyClient = None
try:
    # 고속 다차원 배열 및 벡터 연산을 위한 NumPy 로드 시도
    import numpy as np
except Exception:  # pragma: no cover - numpy is expected to exist, but keep a fallback.
    # NumPy가 없을 경우 None으로 예외 처리
    np = None

# 의미론적 임베딩 모델 객체를 담을 전역 변수 초기화
SentenceTransformer = None
# FAISS 벡터 인덱스 모듈 객체를 담을 전역 변수 초기화
faiss = None


def load_optional_module(module_name):
    # 모듈 사양이 시스템에 존재하는지 확인하고 없으면 None 반환
    if importlib.util.find_spec(module_name) is None:
        return None
    try:
        # 모듈이 존재하면 동적으로 임포트하여 반환
        return importlib.import_module(module_name)
    except Exception:
        # 임포트 도중 예외가 발생하면 안전하게 None 반환
        return None


# UI 필터링 및 질의 분석에 사용되는 지원 대상 8개국 목록
COUNTRY_OPTIONS = ["전체", "대한민국", "일본", "호주", "브라질", "영국", "캐나다", "미국", "중국"]
# 지원 국가별 다양한 동의어 및 영문 표기 매핑 딕셔너리
SUPPORTED_COUNTRY_ALIASES = {
    "대한민국": ["대한민국", "한국", "south korea", "korea"],
    "일본": ["일본", "japan"],
    "호주": ["호주", "australia"],
    "브라질": ["브라질", "brazil"],
    "영국": ["영국", "united kingdom", "uk", "great britain", "britain"],
    "캐나다": ["캐나다", "canada"],
    "미국": ["미국", "united states", "usa", "us", "america"],
    "중국": ["중국", "china"],
}
# 지원하지 않는 국가 목록 (질문 입력 시 가드레일로 차단하기 위한 동의어 매핑)
UNSUPPORTED_COUNTRY_ALIASES = {
    # 아시아 지역 미지원 국가 목록
    "인도": ["인도", "india"], "인도네시아": ["인도네시아", "indonesia"], 
    "파키스탄": ["파키스탄", "pakistan"], "방글라데시": ["방글라데시", "bangladesh"],
    "필리핀": ["필리핀", "philippines"], "베트남": ["베트남", "vietnam"], 
    "터키": ["터키", "turkey", "튀르키예", "turkiye"], "이란": ["이란", "iran"],
    "태국": ["태국", "thailand"], "미얀마": ["미얀마", "myanmar", "burma"], 
    "이라크": ["이라크", "iraq"], "아프가니스탄": ["아프가니스탄", "afghanistan"],
    "사우디아라비아": ["사우디아라비아", "saudi arabia", "saudi"], "우즈베키스탄": ["우즈베키스탄", "uzbekistan"], 
    "말레이시아": ["말레이시아", "malaysia"], "예멘": ["예멘", "yemen"],
    "네팔": ["네팔", "nepal"], "북한": ["북한", "north korea", "dprk"], 
    "스리랑카": ["스리랑카", "sri lanka"], "카자흐스탄": ["카자흐스탄", "kazakhstan"],
    "시리아": ["시리아", "syria"], "캄보디아": ["캄보디아", "cambodia"], 
    "요르단": ["요르단", "jordan"], "아랍에미리트": ["아랍에미리트", "uae", "united arab emirates"],
    "타지키스탄": ["타지키스탄", "tajikistan"], "이스라엘": ["이스라엘", "israel"], 
    "라오스": ["라오스", "laos"], "레바논": ["레바논", "lebanon"],
    "키르기스스탄": ["키르기스스탄", "kyrgyzstan"], "투르크메니스탄": ["투르크메니스탄", "turkmenistan"], 
    "싱가포르": ["싱가포르", "singapore"], "오만": ["오만", "oman"],
    "팔레스타인": ["팔레스타인", "palestine"], "쿠웨이트": ["쿠웨이트", "kuwait"], 
    "조지아": ["조지아", "georgia"], "몽골": ["몽골", "mongolia"],
    "아르메니아": ["아르메니아", "armenia"], "카타르": ["카타르", "qatar"], 
    "바레인": ["바레인", "bahrain"], "동티모르": ["동티모르", "timor-leste", "east timor"],
    "키프로스": ["키프로스", "cyprus"], "부탄": ["부탄", "bhutan"], 
    "몰디브": ["몰디브", "maldives"], "브루나이": ["브루나이", "brunei"],
    "대만": ["대만", "taiwan", "roc"],
    
    # 유럽 지역 미지원 국가 목록
    "러시아": ["러시아", "russia"], "독일": ["독일", "germany"], 
    "프랑스": ["프랑스", "france"], "이탈리아": ["이탈리아", "italy"],
    "스페인": ["스페인", "spain"], "우크라이나": ["우크라이나", "ukraine"], 
    "폴란드": ["폴란드", "poland"], "루마니아": ["루마니아", "romania"],
    "네덜란드": ["네덜란드", "netherlands", "holland"], "벨기에": ["벨기에", "belgium"], 
    "체코": ["체코", "czech", "czechia"], "그리스": ["그리스", "greece"],
    "포르투갈": ["포르투갈", "portugal"], "스웨덴": ["스웨덴", "sweden"], 
    "헝가리": ["헝가리", "hungary"], "벨라루스": ["벨라루스", "belarus"],
    "오스트리아": ["오스트리아", "austria"], "세르비아": ["세르비아", "serbia"], 
    "스위스": ["스위스", "switzerland"], "불가리아": ["불가리아", "bulgaria"],
    "덴마크": ["덴마크", "denmark"], "핀란드": ["핀란드", "finland"], 
    "슬로바키아": ["슬로바키아", "slovakia"], "노르웨이": ["노르웨이", "norway"],
    "아일랜드": ["아일랜드", "ireland"], "크로아티아": ["크로아티아", "croatia"], 
    "몰도바": ["몰도바", "moldova"], "보스니아": ["보스니아", "bosnia", "bosnia and herzegovina"],
    "알바니아": ["알바니아", "albania"], "리투아니아": ["리투아니아", "lithuania"], 
    "북마케도니아": ["북마케도니아", "north macedonia", "macedonia"],
    "슬로베니아": ["슬로베니아", "slovenia"], "라트비아": ["라트비아", "latvia"], 
    "에스토니아": ["에스토니아", "estonia"], "몬테네그로": ["몬테네그로", "montenegro"],
    "룩셈부르크": ["룩셈부르크", "luxembourg"], "몰타": ["몰타", "malta"], 
    "아이슬란드": ["아이슬란드", "iceland"], "안도라": ["안도라", "andorra"],
    "모나코": ["모나코", "monaco"], "리히텐슈타인": ["리히텐슈타인", "liechtenstein"], 
    "산마리노": ["산마리노", "san marino"], "바티칸": ["바티칸", "vatican"],

    # 아메리카 지역 미지원 국가 목록
    "멕시코": ["멕시코", "mexico"], "콜롬비아": ["콜롬비아", "colombia"], 
    "아르헨티나": ["아르헨티나", "argentina"], "페루": ["페루", "peru"],
    "베네수엘라": ["베네수엘라", "venezuela"], "칠레": ["칠레", "chile"], 
    "과테말라": ["과테말라", "guatemala"], "에콰도르": ["에콰도르", "ecuador"],
    "볼리비아": ["볼리비아", "bolivia"], "아이티": ["아이티", "haiti"], 
    "쿠바": ["쿠바", "cuba"], "도미니카 공화국": ["도미니카 공화국", "dominican republic"],
    "온두라스": ["온두라스", "honduras"], "파라과이": ["파라과이", "paraguay"], 
    "엘살바도르": ["엘살바도르", "el salvador"], "니카라과": ["니카라과", "nicaragua"],
    "코스타리카": ["코스타리카", "costa rica"], "파나마": ["파나마", "panama"], 
    "우루과이": ["우루과이", "uruguay"], "자메이카": ["자메이카", "jamaica"],
    "트리니다드 토바고": ["트리니다드 토바고", "trinidad and tobago"], "가이아나": ["가이아나", "guyana"], 
    "수리남": ["수리남", "suriname"], "벨리즈": ["벨리즈", "belize"],
    "바하마": ["바하마", "bahamas"], "바베이도스": ["바베이도스", "barbados"], 
    "세인트루시아": ["세인트루시아", "saint lucia"], "그레나다": ["그레나다", "grenada"],

    # 아프리카 지역 미지원 국가 목록
    "나이지리아": ["나이지리아", "nigeria"], "에티오피아": ["에티오피아", "ethiopia"], 
    "이집트": ["이집트", "egypt"], "민주콩고": ["민주콩고", "dr congo", "congo"],
    "탄자니아": ["탄자니아", "tanzania"], "남아프리카 공화국": ["남아프리카 공화국", "남아공", "south africa"], 
    "케냐": ["케냐", "kenya"], "우간다": ["우간다", "uganda"],
    "수단": ["수단", "sudan"], "알제리": ["알제리", "algeria"], 
    "모로코": ["모로코", "morocco"], "앙골라": ["앙골라", "angola"],
    "모잠비크": ["모잠비크", "mozambique"], "가나": ["가나", "ghana"], 
    "마다가스카르": ["마다가스카르", "madagascar"], "카메룬": ["카메룬", "cameroon"],
    "코트디부아르": ["코트디부아르", "cote d'ivoire", "ivory coast"], "니제르": ["니제르", "niger"], 
    "부르키나파소": ["부르키나파소", "burkina faso"], "말리": ["말리", "mali"],
    "말라위": ["말라위", "malawi"], "잠비아": ["잠비아", "zambia"], 
    "세네갈": ["세네갈", "senegal"], "차드": ["차드", "chad"],
    "소말리아": ["소말리아", "somalia"], "짐바브웨": ["짐바브웨", "zimbabwe"], 
    "기니": ["기니", "guinea"], "르완다": ["르완다", "rwanda"],
    "베냉": ["베냉", "benin"], "부룬디": ["부룬디", "burundi"], 
    "튀니지": ["튀니지", "tunisia"], "남수단": ["남수단", "south sudan"],
    "토고": ["토고", "togo"], "시에라리온": ["시에라리온", "sierra leone"], 
    "리비아": ["리비아", "libya"], "콩고 공화국": ["콩고 공화국", "congo republic"],
    "라이베리아": ["라이베리아", "liberia"], "중앙아프리카 공화국": ["중앙아프리카 공화국", "central african republic"], 
    "모리타니": ["모리타니", "mauritania"], "에리트레아": ["에리트레아", "eritrea"],
    "나미비아": ["나미비아", "namibia"], "감비아": ["감비아", "gambia"], 
    "보츠와나": ["보츠와나", "botswana"], "가봉": ["가봉", "gabon"],
    "레소토": ["레소토", "lesotho"], "기니비사우": ["기니비사우", "guinea-bissau"], 
    "적도 기니": ["적도 기니", "equatorial guinea"], "모리셔스": ["모리셔스", "mauritius"],
    "에스와티니": ["에스와티니", "eswatini", "swaziland"], "지부티": ["지부티", "djibouti"], 
    "코모로": ["코모로", "comoros"], "카보베르데": ["카보베르데", "cape verde"],

    # 오세아니아 지역 미지원 국가 목록
    "파푸아뉴기니": ["파푸아뉴기니", "papua new guinea"], "뉴질랜드": ["뉴질랜드", "new zealand"], 
    "피지": ["피지", "fiji"], "솔로몬 제도": ["솔로몬 제도", "solomon islands"],
    "바누아투": ["바누아투", "vanuatu"], "사모아": ["사모아", "samoa"], 
    "키리바시": ["키리바시", "kiribati"], "미크로네시아": ["미크로네시아", "micronesia"],
    "통가": ["통가", "tonga"], "마셜 제도": ["마셜 제도", "marshall islands"], 
    "팔라우": ["팔라우", "palau"], "투발루": ["투발루", "tuvalu"], 
    "나우루": ["나우루", "nauru"]
}
# 의료기기 인허가 변경 유형 옵션 리스트
CHANGE_TYPE_OPTIONS = [
    "전체",
    "원재료 변경",
    "원재료",
    "원료 변경",
    "라벨 변경",
    "IFU 변경",
    "제조소 변경",
    "제품명 및 모델명 변경",
    "제품명 변경",
    "모델명 변경",
    "유효성 변경",
    "성능 변경",
]

# 캐시 만료 기준 시간 (6시간 = 21,600초)
CACHE_TTL_SECONDS = 60 * 60 * 6
# 메모리에 보관할 최대 캐시 항목 개수
CACHE_MAX_ENTRIES = 128
# 다국어 의미론적 임베딩을 위한 HuggingFace 사전학습 모델 이름
SEMANTIC_MODEL_NAME = "paraphrase-multilingual-MiniLM-L12-v2"

# 임베딩 모델과 벡터 인덱스를 보관할 백엔드 전역 변수 초기화
SEMANTIC_BACKEND = None


def get_secret_value(name, default=""):
    # Streamlit의 secrets 객체 탐색
    secrets = getattr(st, "secrets", None)
    # 기본 반환값 설정
    value = default

    try:
        # secrets가 존재하는 경우 딕셔너리 또는 속성 형태로 값 추출
        if secrets is not None:
            value = secrets.get(name, default) if hasattr(secrets, "get") else secrets[name]
    except Exception:
        # secrets 참조 실패 시 기본값 유지
        value = default

    # 환경변수에 대문자로 등록된 값이 있으면 우선 적용, 없으면 secrets 값 반환
    return os.getenv(name.upper(), value)


# Groq API 키 로드
groq_api_key = get_secret_value("groq_api_key")
# OpenAI API 키 로드
openai_api_key = get_secret_value("openai_api_key")
# Tavily 검색 API 키 로드
tavily_api_key = get_secret_value("tavily_api_key")

# Streamlit 웹페이지 제목, 아이콘, 와이드 레이아웃 설정
st.set_page_config(
    page_title="RA CHATBOT", page_icon="🌐", layout="wide"
)
# 화면 상단 메인 타이틀 출력
st.title("🌐 RA CHATBOT ")

# 사이드바 영역 구성 시작
with st.sidebar:
    # 답변 가능 범위 헤더 출력
    st.markdown("### ✅ 답변 가능 범위")
    # 지원 국가 및 질문 유형 예시 가이드 출력
    st.markdown(
        "**총 8개국:**\n"
        "대한민국(한국), 일본, 호주, 브라질, 영국, 캐나다, 미국, 중국\n\n"
        "**가능한 질문 유형 예시:**\n\n"
        "• 라벨 변경 시 제출 문서\n\n"
        "• 변경 승인 절차 및 요건\n\n"
        "• 규제기관별 지침\n\n"
        "• 변경 분류(Major/Minor) 기준"
    )
    
    # 구분선 추가
    st.markdown("---")
    # 질문 작성 가이드 헤더 출력
    st.markdown("### 📖 효과적인 질문 작성 가이드")
    # 질문에 들어가야 할 필수/권장 키워드와 템플릿 상세 안내
    st.markdown(
        "**✨ 질문에 포함되어야 할 키워드:**\n\n"
        "**1️⃣ 국가명(필수)**\n\n"
        "└ '미국', '일본', '대한민국', '브라질', '호주', '영국', '캐나다', '중국'\n\n"
        "└ 규제기관명(선택): FDA, PMDA, TGA, ANVISA, MHRA, Health Canada, NMPA \n\n\n"
        "**2️⃣ 변경 유형 명시 (권장)**\n\n"
        "└ 가능한 변경 유형: 원재료 변경, 라벨 변경, IFU 변경, 제조소 변경, 제품명 및 모델명 변경\n\n\n"
        "**3️⃣ 구체적인 요청 사항 (필수)**\n\n"
        "└ 제출 문서: '어떤 문서를 제출해야 하나요?', '준비할 서류는?'\n\n"
        "└ 승인 절차: '어떤 절차가 필요한가요?', '심사 기간은?', '단계별 과정은?'\n\n"
        "└ 규정 요건: '요구 사항은 무엇인가요?', '확인해야 할 조건은?'\n\n\n"
        "**4️⃣ 구체적인 상황 설명 (권장)**\n\n"
        "└ 변경의 성격: '라벨의 심볼/로고 변경', '안전 정보 추가', '사용법 변경' 등\n\n"
        "└ 제품 특성: 해당하면 제품 유형 언급\n\n\n"
        "**5️⃣ 한 번에 하나의 질문만 **\n\n"
        "└ ❌ 피해야 할 예시: '미국과 일본의 라벨 변경 차이점을 비교해줘' (비교 기능 미지원)\n\n"
        "└ ✅ 좋은 방법:\n\n"
        "   ① '미국 FDA의 라벨 변경 요건은?'\n\n"
        "   ② '일본 PMDA의 라벨 변경 요건은?'\n\n\n"
        "**📋 최종 질문 작성 템플릿:**\n\n"
        "└ 예: '미국 FDA에서 라벨 변경 시, 제출해야 할 문서는?'\n\n"
        "└ 예: '일본 PMDA의 IFU 변경 승인 절차를 설명해줘'\n\n"
        "└ 예: '브라질에서 제조소 변경에 필요한 요건은?'\n\n\n"
        "**⚠️ 답변 정확도를 위한 팁:**\n\n"
        "• 가능한 한 구체적이고 명확하게 작성하세요\n\n"
        "• 정확하지 않은 답변은 담당 규제 담당자에게 검토를 받아주세요\n\n"
        "• 챗봇이 이해하지 못한 부분은 다시 정의해서 물어봐주세요"
    )


def normalize_text(text):
    # 입력 텍스트가 비어있으면 빈 문자열 반환
    if not text:
        return ""
    # 영문 대문자를 소문자로 변환
    text = str(text).lower()
    # 숫자, 영문, 한글, 공백을 제외한 모든 특수문자를 공백으로 치환
    text = re.sub(r"[^0-9a-z가-힣\s]", " ", text)
    # 연속된 공백을 단일 공백으로 압축하고 양 끝 공백 제거
    text = re.sub(r"\s+", " ", text).strip()
    # 정제된 텍스트 반환
    return text


def tokenize_semantic(text):
    # 텍스트를 정규화
    normalized = normalize_text(text)
    # 정규화된 텍스트가 비어있으면 빈 리스트 반환
    if not normalized:
        return []

    # 토큰을 저장할 리스트 초기화
    tokens = []
    # 띄어쓰기 기준으로 단어 단위 분리
    words = normalized.split()
    # 개별 단어들을 토큰 리스트에 추가
    tokens.extend(words)
    # 인접한 두 단어를 조합한 바이그램(Bigram) 토큰 생성 및 추가
    tokens.extend(f"w:{left}_{right}" for left, right in zip(words, words[1:]))

    # 공백을 제거한 압축 문자열 생성
    compact = normalized.replace(" ", "")
    # 압축 문자열이 존재할 경우 문자 단위 n-gram 생성
    if compact:
        # 2글자(bigram), 3글자(trigram) 단위 문자 시퀀스 토큰 추가
        for size in (2, 3):
            tokens.extend(f"c{size}:{compact[index:index + size]}" for index in range(max(0, len(compact) - size + 1)))

    # 생성된 모든 토큰 리스트 반환
    return tokens


def build_text_vector(text):
    # 토큰화된 리스트의 토큰별 출현 빈도를 Counter 딕셔너리 형태로 반환
    return Counter(tokenize_semantic(text))


def cosine_similarity(left_vector, right_vector):
    # 두 벡터 중 하나라도 비어있으면 유사도 0.0 반환
    if not left_vector or not right_vector:
        return 0.0

    # 연산량 감소를 위해 길이가 더 짧은 벡터를 left_vector로 교환
    if len(left_vector) > len(right_vector):
        left_vector, right_vector = right_vector, left_vector

    # 두 벡터의 내적(Dot Product) 계산
    dot = sum(value * right_vector.get(token, 0) for token, value in left_vector.items())
    # 좌측 벡터의 크기(L2 Norm) 계산
    left_norm = math.sqrt(sum(value * value for value in left_vector.values()))
    # 우측 벡터의 크기(L2 Norm) 계산
    right_norm = math.sqrt(sum(value * value for value in right_vector.values()))
    # 벡터 크기가 0이면 유사도 0.0 반환
    if not left_norm or not right_norm:
        return 0.0
    # 코사인 유사도 공식(내적 / 크기곱) 결과 반환
    return dot / (left_norm * right_norm)


def normalize_vector(vector):
    # numpy가 없거나 입력 벡터가 없으면 그대로 반환
    if np is None or vector is None:
        return vector

    # 벡터를 float32 타입의 NumPy 배열로 변환
    vector = np.asarray(vector, dtype="float32")
    # 벡터의 유클리드 노름(크기) 계산
    norm = np.linalg.norm(vector)
    # 노름이 0이면 그대로 반환
    if not norm:
        return vector
    # 단위 벡터로 정규화하여 반환
    return vector / norm


def get_semantic_model():
    # 전역 변수 SEMANTIC_BACKEND 참조
    global SEMANTIC_BACKEND

    # 이미 로드된 모델이 있다면 재사용
    if SEMANTIC_BACKEND and SEMANTIC_BACKEND.get("model") is not None:
        return SEMANTIC_BACKEND["model"]

    # sentence_transformers 모듈 동적 로드
    sentence_transformers_module = load_optional_module("sentence_transformers")
    # 모듈이 설치되어 있지 않으면 None 반환
    if sentence_transformers_module is None:
        return None

    try:
        # 사전학습된 SentenceTransformer 임베딩 모델 인스턴스화
        model = sentence_transformers_module.SentenceTransformer(SEMANTIC_MODEL_NAME)
    except Exception:
        # 모델 로드 실패 시 None 반환
        return None

    # 전역 백엔드 딕셔너리가 없으면 초기화
    if SEMANTIC_BACKEND is None:
        SEMANTIC_BACKEND = {}
    # 백엔드에 로드된 모델 등록
    SEMANTIC_BACKEND["model"] = model
    # 모델 객체 반환
    return model


def build_semantic_backend(documents):
    # 임베딩 모델 인스턴스 가져오기
    model = get_semantic_model()
    # 모델이나 NumPy가 없으면 빌드 불가하므로 None 반환
    if model is None or np is None:
        return None

    # 문서 리스트에서 임베딩용 텍스트 필드 추출
    texts = [doc.get("semantic_text", doc.get("content", "")) for doc in documents]
    # 텍스트가 없으면 None 반환
    if not texts:
        return None

    try:
        # 모든 문서 텍스트를 벡터 임베딩으로 일괄 변환
        embeddings = model.encode(texts, convert_to_numpy=True, show_progress_bar=False)
    except Exception:
        # 임베딩 생성 실패 시 None 반환
        return None

    # 임베딩을 float32 NumPy 배열로 변환
    embeddings = np.asarray(embeddings, dtype="float32")
    # 배열 차원이 2차원이 아니거나 비어있으면 None 반환
    if embeddings.ndim != 2 or embeddings.size == 0:
        return None

    # 모든 임베딩 벡터를 단위 벡터로 정규화
    embeddings = np.vstack([normalize_vector(vector) for vector in embeddings])

    # 백엔드 객체 구성
    backend = {
        "model": model,
        "embeddings": embeddings,
    }

    # faiss 모듈 동적 로드 시도
    faiss_module = load_optional_module("faiss")
    if faiss_module is not None:
        try:
            # 내적(Inner Product) 기반의 고속 FAISS 인덱스 생성
            index = faiss_module.IndexFlatIP(int(embeddings.shape[1]))
            # 인덱스에 문서 임베딩 벡터 등록
            index.add(embeddings)
            # 백엔드에 FAISS 인덱스 저장
            backend["index"] = index
        except Exception:
            # 실패 시 인덱스를 None으로 설정
            backend["index"] = None

    # 완성된 백엔드 딕셔너리 반환
    return backend


def semantic_rank_documents(query, documents, backend):
    # 백엔드나 문서 리스트가 없으면 빈 리스트 반환
    if not backend or not documents:
        return []

    # 백엔드에서 모델과 임베딩 배열 추출
    model = backend.get("model")
    embeddings = backend.get("embeddings")
    # 필수 구성 요소가 누락되었으면 빈 리스트 반환
    if model is None or embeddings is None or np is None:
        return []

    try:
        # 사용자 질문을 임베딩 벡터로 변환
        query_vector = model.encode([query], convert_to_numpy=True, show_progress_bar=False)[0]
    except Exception:
        # 인코딩 실패 시 빈 리스트 반환
        return []

    # 질문 벡터 정규화
    query_vector = normalize_vector(query_vector)
    # 정규화 실패 시 빈 리스트 반환
    if query_vector is None:
        return []

    # FAISS 모듈 로드 시도
    faiss_module = load_optional_module("faiss")
    # FAISS 인덱스가 준비되어 있다면 FAISS 기반 고속 검색 실행
    if backend.get("index") is not None and faiss_module is not None:
        try:
            # 질문 벡터와 유사한 상위 문서(최대 12개) 인덱스 및 유사도 점수 검색
            scores, indices = backend["index"].search(np.asarray([query_vector], dtype="float32"), min(12, len(documents)))
            ranked = []
            # 검색 결과를 순회하며 유효한 인덱스의 문서를 랭킹 리스트에 추가
            for score, idx in zip(scores[0], indices[0]):
                if idx < 0:
                    continue
                ranked.append((float(score), documents[int(idx)]))
            # 랭킹 결과 반환
            return ranked
        except Exception:
            # FAISS 검색 실패 시 행렬곱 연산 방식으로 폴백
            pass

    # FAISS가 없을 경우 직접 행렬 곱(@)으로 모든 문서와의 코사인 유사도 계산
    scores = embeddings @ query_vector
    # 각 점수와 문서를 튜플로 묶어 리스트 생성
    ranked = [(float(score), documents[idx]) for idx, score in enumerate(scores)]
    # 유사도 점수 기준 내림차순 정렬
    ranked.sort(key=lambda item: item[0], reverse=True)
    # 정렬된 문서 랭킹 리스트 반환
    return ranked


def chunk_text(text, max_length=650):
    # 텍스트를 문자열로 변환하고 양 끝 공백 제거
    clean_text = str(text or "").strip()
    # 텍스트가 비어있으면 빈 리스트 반환
    if not clean_text:
        return []

    # 파이프(|) 구분자를 기준으로 텍스트 분할
    parts = [part.strip() for part in re.split(r"\s*\|\s*", clean_text) if part.strip()]
    # 파이프가 없어서 분할되지 않은 경우 문장 종결 부호 기준으로 재분할
    if len(parts) == 1:
        parts = [part.strip() for part in re.split(r"(?<=[。.!?])\s+", clean_text) if part.strip()]

    # 최종 분할 청크들을 담을 리스트
    chunks = []
    # 현재 청크 버퍼
    current = ""
    # 각 조각들을 순회하며 최대 길이에 맞게 결합
    for part in parts:
        # 버퍼가 비어있으면 첫 조각 할당
        if not current:
            current = part
            continue

        # 기존 청크에 새 조각을 덧붙인 후보 문자열 생성
        candidate = f"{current} | {part}"
        # 최대 길이를 넘지 않으면 결합 유지
        if len(candidate) <= max_length:
            current = candidate
        else:
            # 최대 길이를 초과하면 이전 청크를 저장하고 새로운 청크 시작
            chunks.append(current)
            current = part

    # 마지막으로 남은 버퍼 청크 추가
    if current:
        chunks.append(current)

    # 청크 리스트 반환 (없으면 원본 텍스트 리스트 반환)
    return chunks or [clean_text]


def build_conversation_memory(messages, max_turns=3):
    # 최근 턴의 대화 문자열들을 담을 리스트
    turns = []
    # 최신 메시지부터 역순으로 탐색
    for message in reversed(messages):
        role = message.get("role")
        content = str(message.get("content", "")).strip()
        # user 또는 assistant 역할이 아니거나 내용이 없으면 건너뜀
        if role not in {"user", "assistant"} or not content:
            continue
        # 발화자 역할과 내용을 포맷팅하여 추가
        turns.append(f"{role}: {content}")
        # 설정된 최대 턴 수(user/assistant 쌍 고려하여 * 2)에 도달하면 중단
        if len(turns) >= max_turns * 2:
            break

    # 수집된 대화가 없으면 빈 문자열 반환
    if not turns:
        return ""

    # 역순으로 수집했으므로 다시 시간 순서대로 뒤집어 줄바꿈 문자로 연결
    return "\n".join(reversed(turns))

def search_external_regulatory_info(query):
    # Tavily 라이브러리 부재 시 안내 메시지 반환
    if TavilyClient is None:
        return "⚠️ Tavily 패키지가 설치되지 않아 외부 웹 검색이 생략되었습니다."
    # API 키 부재 시 안내 메시지 반환
    if not tavily_api_key:
        return "⚠️ Tavily API 키가 설정되지 않아 외부 웹 검색이 생략되었습니다."
    try:
        # Tavily 검색 클라이언트 객체 생성
        client = TavilyClient(api_key=tavily_api_key)
        
        # 검색 대상 도메인을 의료기기 규제로 한정하기 위한 쿼리 보강
        domain_query = f"의료기기 인허가 규제 {query}"
        
        # 외부 웹 검색 실행 (속도 최적화 및 토큰 절약을 위해 basic 모드, 상위 3건만 요청)
        response = client.search(
            query=domain_query,
            search_depth="basic", # advanced 보다 속도가 빠르고 토큰 소모가 적음
            max_results=3,        # 컨텍스트 길이(토큰)를 줄이기 위해 상위 3개만 추출
            include_answer=False
        )
        
        # 검색 결과 리스트 추출
        results = response.get("results", [])
        # 결과가 없으면 빈 문자열 반환
        if not results:
            return ""
        
        # LLM 프롬프트에 주입할 컨텍스트 헤더 구성
        search_context = "\n[🌐 외부 웹 검색 결과 (최신 동향)]\n"
        # 검색된 각 웹페이지의 제목, 내용 요약, URL을 번호 매겨 누적
        for i, res in enumerate(results, 1):
            search_context += f"{i}. 제목: {res['title']}\n   내용: {res['content']}\n   출처: <{res['url']}>\n"
        # 완성된 외부 검색 컨텍스트 문자열 반환
        return search_context
    except Exception as e:
        # 검색 도중 예외 발생 시 에러 메시지 반환
        return f"외부 검색 엔진 오류: {e}"

def rewrite_query_for_search(query, selected_countries=None, selected_change_types=None):
    # 쿼리를 문자열로 변환하고 양 끝 공백 제거
    rewritten = str(query or "").strip()
    # 구어체 및 약어를 표준 검색어로 매핑하기 위한 사전
    replacements = {
        "어케": "어떻게",
        "어떡": "어떻게",
        "라벨링": "라벨",
        "스텐트": "스텐트",
        "ifu": "IFU",
    }

    # 쿼리 정규화
    normalized = normalize_text(rewritten)
    # 구어체/약어 치환 규칙 적용
    for source, target in replacements.items():
        if source in normalized:
            rewritten = re.sub(source, target, rewritten, flags=re.IGNORECASE)
            normalized = normalize_text(rewritten)

    # 사이드바에서 선택된 국가가 있다면 검색어 앞에 국가명 힌트 추가
    if selected_countries:
        country_hint = " / ".join(selected_countries)
        if country_hint and country_hint not in rewritten:
            rewritten = f"{country_hint} {rewritten}".strip()

    # 사이드바에서 선택된 변경 유형이 있다면 검색어 뒤에 변경 유형 힌트 추가
    if selected_change_types:
        change_hint = " / ".join(selected_change_types)
        if change_hint and change_hint not in rewritten:
            rewritten = f"{rewritten} {change_hint}".strip()

    # '라벨' 관련 질문인데 제출 키워드가 없으면 '제출 문서' 키워드 보강
    if "라벨" in normalize_text(rewritten) and "제출" not in normalize_text(rewritten):
        rewritten = f"{rewritten} 제출 문서"

    # '변경' 관련 질문인데 승인 키워드가 없으면 '승인 요건' 키워드 보강
    if "변경" in normalize_text(rewritten) and "승인" not in normalize_text(rewritten):
        rewritten = f"{rewritten} 승인 요건"

    # 연속된 공백을 단일 공백으로 치환 후 반환
    return re.sub(r"\s+", " ", rewritten).strip()


def get_query_filters(query):
    # 질문 정규화
    query_clean = normalize_text(query)
    # "전체"를 제외한 실제 국가 목록 정의
    target_countries = COUNTRY_OPTIONS[1:]
    # 질문에 직접 언급된 대상 국가 추출
    found_countries = [country for country in target_countries if country in query]

    # 매칭할 정확한 변경 유형 키워드 목록
    exact_change_types = [
        "원재료 변경",
        "원재료",
        "원료 변경",
        "라벨 변경",
        "ifu 변경",
        "제조소 변경",
        "제품명 및 모델명 변경",
        "제품명 변경",
        "모델명 변경",
        "유효성 변경",
        "성능 변경",
    ]
    # 질문 내에 존재하는 첫 번째 변경 유형 키워드 추출 (없으면 빈 문자열)
    requested_change_type = next((change_type for change_type in exact_change_types if change_type in query_clean), "")

    # 추출된 국가 목록과 변경 유형 반환
    return found_countries, requested_change_type


def get_query_source(query):
    # 질문 정규화
    query_clean = normalize_text(query)
    # 변경 규정 관련 시그널 키워드 목록
    change_signals = [
        "변경",
        "라벨",
        "원재료",
        "원료",
        "제조소",
        "제품명",
        "모델명",
        "ifu",
        "심사기간",
        "일부변경",
        "경미변경",
        "major",
        "minor",
        "approval change",
    ]
    # 시그널 단어가 포함되어 있으면 "change"(변경 규정) 소스로 분류
    if any(signal in query_clean for signal in change_signals):
        return "change"

    # 일반 규정 관련 시그널 키워드 목록
    general_signals = ["규정", "허가", "승인", "확인", "방법", "어떻게", "기준", "요건", "절차"]
    # 일반 시그널이 포함되어 있으면 "general"(일반 규정) 소스로 분류
    if any(signal in query_clean for signal in general_signals):
        return "general"

    # 기본값으로 "general" 반환
    return "general"


def _normalize_country_alias(alias):
    # 국가 별칭에서 공백을 완전히 제거하여 정규화
    return normalize_text(alias).replace(" ", "")


def extract_country_mentions(question):
    # 질문에서 공백 제거 및 정규화
    normalized_question = normalize_text(question).replace(" ", "")
    # 감지된 지원 국가 리스트
    supported_matches = []
    # 감지된 미지원 국가 리스트
    unsupported_matches = []

    # 지원 국가 동의어 대조
    for country, aliases in SUPPORTED_COUNTRY_ALIASES.items():
        if any(_normalize_country_alias(alias) in normalized_question for alias in aliases):
            supported_matches.append(country)

    # 미지원 국가 동의어 대조
    for country, aliases in UNSUPPORTED_COUNTRY_ALIASES.items():
        if any(_normalize_country_alias(alias) in normalized_question for alias in aliases):
            unsupported_matches.append(country)

    # 각각 매칭된 국가 리스트 반환
    return supported_matches, unsupported_matches


def should_block_country_question(question):
    # 질문 내 지원/미지원 국가 언급 추출
    supported_matches, unsupported_matches = extract_country_mentions(question)
    # 미지원 국가가 하나라도 포함되어 있으면 차단(True) 및 해당 국가명 반환
    if unsupported_matches:
        return True, unsupported_matches
    # 차단하지 않음(False)과 지원 국가 목록 반환
    return False, supported_matches


def build_country_scope_guard_response(unsupported_countries):
    # 차단된 국가명을 쉼표로 연결
    blocked_label = ", ".join(unsupported_countries) if unsupported_countries else "해당 국가"
    # 미지원 국가 안내 표준 템플릿 응답 반환
    return (
        f"현재 {blocked_label}를 포함한 8개국 외의 답변은 어렵습니다. 공식 사이트를 참조하세요.\n\n"
        "지원 국가: 대한민국, 일본, 호주, 브라질, 영국, 캐나다, 미국, 중국"
    )


def is_global_regulatory_query(question):
    # 질문 정규화
    text = normalize_text(question)
    # 빈 문자열이면 False 반환
    if not text:
        return False

    # 의료기기 규제 도메인 식별 키워드 리스트
    allowed_markers = [
        "규제",
        "규정",
        "허가",
        "승인",
        "라벨",
        "제출 문서",
        "제출",
        "변경",
        "pmda",
        "fda",
        "mah",
        "meddev",
        "의료기기",
        "국가",
        "기관",
        "문서",
        "심사",
        "요건",
        "절차",
        "기준",
        "가이드라인",
        "guideline",
        "규제기관",
        "국가별",
        "국가별 규정",
        "국가별 규제",
        "미국",
        "일본",
        "브라질",
        "캐나다",
        "호주",
        "영국",
        "대한민국",
    ]
    # 도메인 키워드가 하나라도 있으면 유효한 질의로 판정
    if any(marker in text for marker in allowed_markers):
        return True

    # 일상 대화 등 도메인 외 질문인 경우 False 반환
    return False


def build_scope_guard_response(question):
    # 도메인 이탈 질문 시 출력할 표준 차단 메시지 반환
    return (
        "현재 이 어시스턴트는 글로벌 의료기기 규제 정보에 한정해서 답변합니다.\n\n"
        "질문이 글로벌 의료기기 규제/규정/허가/라벨/제출 문서/변경 승인 절차와 관련된 경우에만 답변할 수 있습니다.\n"
        "그 외 주제(예: 일반 일상, 일반 지식, 비규제/비의료기기 주제)는 답변 범위를 벗어납니다."
    )


def normalize_filter_values(selected_values):
    # 선택된 값이 없으면 빈 리스트 반환
    if not selected_values:
        return []
    # '전체' 옵션과 빈 값을 제외한 실제 필터 항목만 추출하여 반환
    return [value for value in selected_values if value and value != "전체"]


def cleanup_cache(cache_store):
    # 현재 시간 측정
    now = time.time()
    # TTL(유효기간)이 초과된 만료된 캐시 키 목록 식별
    expired_keys = [key for key, payload in cache_store.items() if now - payload.get("created_at", 0) > CACHE_TTL_SECONDS]
    # 만료된 캐시 항목 제거
    for key in expired_keys:
        cache_store.pop(key, None)

    # 캐시 항목 수가 최대 허용치를 초과하는 경우 가장 오래된 항목부터 삭제
    while len(cache_store) > CACHE_MAX_ENTRIES:
        oldest_key = min(cache_store.items(), key=lambda item: item[1].get("created_at", 0))[0]
        cache_store.pop(oldest_key, None)


def get_cached_answer(cache_store, cache_key):
    # 캐시 정리 작업 선행 실행
    cleanup_cache(cache_store)
    # 키에 해당하는 캐시 데이터 조회
    payload = cache_store.get(cache_key)
    # 데이터가 없으면 None 반환
    if not payload:
        return None
    # 조회된 데이터가 TTL을 초과했으면 삭제하고 None 반환
    if time.time() - payload.get("created_at", 0) > CACHE_TTL_SECONDS:
        cache_store.pop(cache_key, None)
        return None
    # 유효한 캐시 답변 텍스트 반환
    return payload.get("answer")


def set_cached_answer(cache_store, cache_key, answer):
    # 캐시 딕셔너리에 답변 내용과 생성 시간 저장
    cache_store[cache_key] = {
        "answer": answer,
        "created_at": time.time(),
    }
    # 저장 후 캐시 크기 정리
    cleanup_cache(cache_store)


@st.cache_data
def load_documents():
    # 전체 로드된 문서를 담을 리스트
    documents = []

    def first_present(row, *columns):
        # 여러 컬럼 후보 중 첫 번째로 유효한(비어있지 않은) 값을 반환하는 내부 함수
        for column in columns:
            value = row.get(column, "")
            if pd.notna(value) and str(value).strip():
                return str(value).strip()
        return ""

    def add_document(source, content, **metadata):
        # 본문 내용을 청크 단위로 분할
        chunks = chunk_text(content)
        # 분할된 각 청크를 순회하며 메타데이터와 벡터를 생성하여 등록
        for index, chunk in enumerate(chunks, 1):
            # 의미론적 임베딩을 위한 국가명, 변경유형 결합 텍스트 생성
            semantic_text_parts = [str(metadata.get("country", "")).strip(), str(metadata.get("change_type", "")).strip(), chunk]
            document = {
                "source": source,
                "content": chunk,
                "semantic_text": " ".join(part for part in semantic_text_parts if part),
                "normalized_content": normalize_text(chunk),
                "vector": build_text_vector(chunk),
                "chunk_index": index,
                "chunk_total": len(chunks),
            }
            # 추가 메타데이터 병합
            document.update(metadata)
            # 문서 리스트에 추가
            documents.append(document)

    # 1. 일반 규정 엑셀 파일 로드 및 문서화
    if os.path.exists("general regulations.xlsx"):
        df_reg = pd.read_excel("general regulations.xlsx")
        for _, row in df_reg.iterrows():
            text = (
                f"[{row.get('국가', '')} {row.get('규제기관명', '')}] "
                f"일반규정: {first_present(row, '관련 법령', '관련 법령/규정')} / "
                f"규정 확인 사이트: {first_present(row, '규정 확인 사이트', '규제기관 홈페이지')} / "
                f"참고: {first_present(row, '공유 파일 주소', '참고')}"
            )
            add_document(
                "general",
                text,
                country=str(row.get("국가", "")).strip(),
                regulation_agency=str(row.get("규제기관명", "")).strip(),
            )

    # 2. 변경 규정 엑셀 파일 로드 및 문서화
    if os.path.exists("change regulations.xlsx"):
        df_change = pd.read_excel("change regulations.xlsx")
        for _, row in df_change.iterrows():
            change_type = str(row.get('변경 유형', '')).strip()
            text = (
                f"[{row.get('국가', '')} 변경규정] 유형: {row.get('변경 유형', '')} "
                f"({row.get('변경 분류 (Major/Minor 등)', '')}) | 제출문서: {row.get('제출 문서', '')} "
                f"| 심사기간: {first_present(row, '심사 기간 (영업일 기준)', '심사 기간')} "
                f"| 변경허가 필요 여부: {first_present(row, '변경허가 필요 여부', '허가 필요 여부')} "
                f"| 관련 규정: {first_present(row, '관련 규정', '관련 법령/규정')} "
                f"| 조항 위치: {first_present(row, '조항 위치', '근거 문서 위치')} "
                f"| 근거 원문: {first_present(row, '근거 원문', '근거 원문 및 국문 번역')} "
                f"| 국문 번역: {first_present(row, '근거 원문 국문 번역', '근거 원문 및 국문 번역')}"
            )
            add_document(
                "change",
                text,
                country=str(row.get("국가", "")).strip(),
                change_type=change_type,
                change_class=str(row.get("변경 분류 (Major/Minor 등)", "")).strip(),
            )

    # 3. 추가 컨텍스트 엑셀 파일 로드 및 문서화
    if os.path.exists("additional regulations.xlsx"):
        df_add = pd.read_excel("additional regulations.xlsx")
        for _, row in df_add.iterrows():
            text = row.get("AI_Search_Context") or ""
            if text:
                add_document(
                    "additional",
                    text,
                    country=str(row.get("국가", "")).strip(),
                    change_type=str(row.get("변경 유형", "")).strip(),
                )

    # 구축된 전체 지식베이스 문서 리스트 반환
    return documents


def sanitize_urls_for_markdown(text):
    # 텍스트가 없으면 원본 그대로 반환
    if not text:
        return text

    # HTTP/HTTPS URL 탐색을 위한 정규표현식 컴파일
    url_pattern = re.compile(r"https?://[A-Za-z0-9\-._~:/?#@!$&'()+,;=%]+")

    def replace_url(match):
        # URL 앞뒤에 꺾쇠괄호(<, >)를 추가하여 마크다운 링크 파싱 에러 방지
        url = match.group(0)
        return f"<{url}>"

    # 변환된 텍스트 반환
    return url_pattern.sub(replace_url, text)


def get_active_filters(query, selected_country=None, selected_change_type=None):
    # 질문 텍스트에서 국가 및 변경 유형 필터 추출
    query_countries, query_change_type = get_query_filters(query)

    # 사이드바에서 선택된 값 정규화
    selected_country_values = normalize_filter_values(selected_country)
    selected_change_type_values = normalize_filter_values(selected_change_type)

    # 사이드바에서 선택된 국가가 있다면 질문 추출값 대신 우선 적용
    if selected_country_values:
        query_countries = selected_country_values

    # 사이드바에서 선택된 변경 유형이 있다면 우선 적용, 없으면 질문 추출값 사용
    if selected_change_type_values:
        query_change_type = selected_change_type_values
    elif query_change_type:
        query_change_type = [query_change_type]
    else:
        query_change_type = []

    # 최종 결정된 국가 및 변경 유형 필터 리스트 반환
    return query_countries, query_change_type


def retrieve_documents(query, documents, top_k=3, selected_country=None, selected_change_type=None):
    # 참조할 문서가 없으면 빈 리스트 반환
    if not documents:
        return []

    # 사이드바 선택 필터 값 정규화
    selected_country_values = normalize_filter_values(selected_country)
    selected_change_type_values = normalize_filter_values(selected_change_type)

    # 검색에 최적화된 형태로 질문 재작성
    rewritten_query = rewrite_query_for_search(query, selected_country_values, selected_change_type_values)
    # 재작성된 검색어 정규화
    query_clean = normalize_text(rewritten_query)
    # 검색어의 BoW 단어 벡터 생성
    query_vector = build_text_vector(rewritten_query)
    # 필터 및 소스 유형 도출
    found_countries, requested_change_types = get_active_filters(query, selected_country_values, selected_change_type_values)
    requested_source = get_query_source(rewritten_query)
    # 2글자 이상인 주요 검색 키워드 추출
    keywords = [kw for kw in query_clean.split() if len(kw) >= 2]

    # 의미론적 임베딩 기반 상위 문서 랭킹 수행
    semantic_ranked = semantic_rank_documents(rewritten_query, documents, SEMANTIC_BACKEND)
    if semantic_ranked:
        ranked_docs = semantic_ranked
    else:
        # 임베딩 백엔드가 없을 경우 코사인 유사도 기반으로 계산
        ranked_docs = []
        for doc in documents:
            content_clean = doc.get("normalized_content") or normalize_text(doc["content"])
            score = cosine_similarity(query_vector, doc.get("vector") or build_text_vector(doc["content"])) * 100

            # 질문 소스와 문서 소스가 일치하면 가산점 부여
            if doc.get("source") == requested_source:
                score += 8

            # 추가 규정 문서일 경우 기본 가산점 부여
            if doc.get("source") == "additional":
                score += 1

            ranked_docs.append((score, doc))

        # 유사도 기준 내림차순 정렬
        ranked_docs.sort(key=lambda x: x[0], reverse=True)

    # 필터 및 가중치 적용을 위한 재채점 리스트
    scored = []
    for score, doc in ranked_docs:
        # 타겟 소스나 추가 규정이 아니면 제외
        if doc.get("source") not in {requested_source, "additional"}:
            continue

        content_clean = doc.get("normalized_content") or normalize_text(doc["content"])
        change_type_clean = normalize_text(str(doc.get("change_type", "")))

        # 질문에 지정된 국가가 본문에 포함되어 있지 않으면 제외
        if found_countries and not any(country in doc["content"] for country in found_countries):
            continue

        # 변경 규정 문서의 경우 지정된 변경 유형이 매칭되지 않으면 제외
        if requested_change_types and doc.get("source") == "change":
            if not any(change_type in change_type_clean for change_type in requested_change_types):
                continue

        # 해당 국가명이 포함되어 있으면 높은 가산점(+12) 부여
        for country in found_countries:
            if country in doc["content"]:
                score += 12

        # 검색 키워드가 포함되어 있을 때마다 가산점(+2) 부여
        for kw in keywords:
            if kw in content_clean:
                score += 2

        # 특정 도메인 핵심어 매칭 가산점
        if "라벨" in query_clean and "라벨" in content_clean:
            score += 3
        if "문서" in query_clean and "문서" in content_clean:
            score += 2
        if "변경" in query_clean and "변경" in content_clean:
            score += 2

        scored.append((score, doc))

    # 최종 점수 기준 내림차순 정렬
    scored.sort(key=lambda x: x[0], reverse=True)
    # 점수가 0점보다 큰 유효 문서만 필터링
    filtered_docs = [doc for score, doc in scored if score > 0]

    # 변경 유형이 지정된 경우 정확히 일치하는 문서를 우선 추출하여 반환
    if requested_change_types:
        exact_docs = [
            doc for doc in filtered_docs
            if any(change_type in normalize_text(str(doc.get("change_type", ""))) for change_type in requested_change_types)
        ]
        if exact_docs:
            return exact_docs[:top_k]

    # 상위 top_k개 문서 반환
    return filtered_docs[:top_k]


def build_retrieval_context(query, docs, conversation_memory="", rewritten_query=""):
    # 이전 대화 기억이 있으면 컨텍스트 상단에 추가
    if conversation_memory:
        lines = ["최근 대화 기억:", conversation_memory, ""]
    else:
        lines = []

    # 원본 질문 추가
    lines.append(f"원문 질문: {query}")

    # 재작성된 검색 질문이 있으면 추가
    if rewritten_query:
        lines.append(f"검색용 질문: {rewritten_query}")

    # 참조 문서가 없으면 메시지 추가 후 반환
    if not docs:
        lines.append("참조 문서가 없습니다.")
        return "\n".join(lines)

    # 검색된 상위 3개 문서 포맷팅 추가
    lines.append("참조 문서:")
    for index, doc in enumerate(docs[:3], 1):
        content = str(doc.get("content", "")).strip()
        metadata = []
        # 메타데이터(국가, 변경유형) 추출
        if doc.get("country"):
            metadata.append(str(doc.get("country")))
        if doc.get("change_type"):
            metadata.append(str(doc.get("change_type")))
        # 메타데이터가 있으면 괄호 표기 후 내용 결합
        if metadata:
            lines.append(f"[{index}] ({' / '.join(metadata)}) {content}")
        else:
            lines.append(f"[{index}] {content}")
    # 개행 문자로 연결하여 최종 컨텍스트 반환
    return "\n".join(lines)


def is_simple_question(question):
    # 질문 정규화
    text = normalize_text(question)
    # 질문이 비어있으면 단순 조회로 간주
    if not text:
        return True
    
    # 비교/분석 등 복합 추론이 필요한 키워드 정의
    comparison_markers = ["어떻게 다른", "차이", "vs", "차이점", "비교", "분석", "특징", "추세", "개선", "방안", "중요", "다른가요"]
    # 비교 키워드가 포함되어 있으면 단순 조회가 아님(False)
    if any(marker in text for marker in comparison_markers):
        return False  
    
    # 단순 정보 조회를 의미하는 키워드 정의
    info_markers = ["어떤 문서", "제출 문서", "제출해야", "준비해야", "어떤 규정", "무엇", "뭐"]
    # 단순 정보 조회 키워드 포함 여부 반환
    return any(marker in text for marker in info_markers)


def build_file_based_answer(query, docs):
    # 문서가 없을 때의 기본 안내 반환
    if not docs:
        return "관련 문서를 찾지 못했습니다. 질문을 더 구체적으로 써주시면 더 정확하게 도와드릴 수 있습니다."

    # 로컬 파싱 요약 답변 헤더
    answer = "**관련 규정 정보 (자체 검색 요약):**\n\n"
    
    # 상위 3개 문서를 순회하며 필드 단위 파싱
    for index, doc in enumerate(docs[:3], 1):
        content = doc['content']
        # 대한민국 규정 문서 여부 확인
        is_korean_document = content.startswith("[대한민국")
        # 파이프(|) 단위로 분리
        parts = [part.strip() for part in content.split('|')]
        answer += f"{index}. {parts[0]}\n"
        
        # 세부 필드 순회
        for part in parts[1:]:
            if part:
                # 대한민국 문서는 중복 방지를 위해 국문 번역 항목 생략
                if is_korean_document and part.startswith("국문 번역:"):
                    continue
                # 콜론(:)이 포함되어 있으면 항목명과 내용으로 분리하여 들여쓰기 출력
                if ':' in part:
                    field, value = part.split(':', 1)
                    answer += f"   - {field.strip()}: {value.strip()}\n"
                else:
                    answer += f"   - {part}\n"
        answer += "\n"
    
    # 주의사항 및 법적 면책 안내 추가
    answer += (
        "**정리 및 주의사항:**\n"
        "위의 문서들을 바탕으로 다음을 확인하세요:\n\n"
        "• 해당 국가/규제기관의 정확 요건\n"
        "• 필요한 제출 문서\n"
        "• 변경 분류 및 승인 절차\n\n"
        "⚠️ 이는 내부 문서 검색 결과이므로 정확한 해석은 담당 규제 담당자에게 검토를 받아주세요."
    )
    return answer


def append_evidence_translation(answer, docs):
    # 번역 블록들을 담을 리스트
    translation_blocks = []

    # 검색된 상위 3개 문서를 순회
    for doc in docs[:3]:
        content = str(doc.get("content", ""))
        # 대한민국 문서는 번역 블록 생성이 불필요하므로 건너뜀
        if content.startswith("[대한민국"):
            continue

        lines = []
        # 각 필드 중 근거 원문 및 국문 번역 라인만 추출
        for part in content.split("|"):
            part = part.strip()
            if part.startswith("근거 원문:"):
                lines.append(f"   - {part}")
            elif part.startswith("국문 번역:"):
                lines.append(f"   - {part}")

        # 추출된 라인이 있으면 블록으로 추가
        if lines:
            translation_blocks.append("\n".join(lines))

    # 번역 데이터가 없으면 기존 답변 그대로 반환
    if not translation_blocks:
        return answer

    # 답변 내용에 이미 국문 번역이 포함되어 있으면 중복 방지를 위해 그대로 반환
    if "국문 번역" in answer:
        return answer

    # 답변 하단에 근거 원문 및 국문 번역 섹션을 결합하여 반환
    return (
        answer
        + "\n\n**근거 원문 및 국문 번역:**\n"
        + "\n\n".join(translation_blocks)
    )


def finalize_answer_text(answer, docs):
    # 근거 원문 및 번역 텍스트를 첨부
    finalized = append_evidence_translation(answer, docs)
    # URL 링크를 마크다운에 적합하게 정제 후 반환
    return sanitize_urls_for_markdown(finalized)


def build_cache_key(query, selected_country, selected_change_type, selected_docs, is_simple):
    # 선택된 국가 및 변경 유형 정규화
    selected_countries = normalize_filter_values(selected_country)
    selected_change_types = normalize_filter_values(selected_change_type)
    # 검색된 상위 3개 문서의 고유 시그니처 문자열 생성
    doc_signature = "|".join(
        f"{doc.get('source', '')}:{doc.get('country', '')}:{doc.get('change_type', '')}:{doc.get('chunk_index', 0)}"
        for doc in selected_docs[:3]
    )
    # 모든 조건을 조합한 고유 캐시 키 생성 및 반환
    return "||".join(
        [
            normalize_text(query),
            ",".join(selected_countries) or "전체",
            ",".join(selected_change_types) or "전체",
            doc_signature,
            "simple" if is_simple else "complex",
        ]
    )


def parse_stream_content(raw_line):
    # 빈 데이터 라인이면 None 반환
    if not raw_line:
        return None

    # 바이트 스트림일 경우 UTF-8 문자열로 디코딩
    if isinstance(raw_line, bytes):
        raw_line = raw_line.decode("utf-8", errors="replace")

    # Server-Sent Events의 "data: " 접두사 확인
    if not raw_line.startswith("data: "):
        return None

    # 접두사를 제거한 JSON 페이로드 추출
    payload = raw_line[6:].strip()
    # 스트림 종료 시그널 확인
    if payload == "[DONE]":
        return "[DONE]"

    try:
        # JSON 파싱
        chunk = json.loads(payload)
    except json.JSONDecodeError:
        return None

    # 응답 choices 배열 확인
    choices = chunk.get("choices") or []
    if not choices:
        return None

    # 스트리밍 텍스트 조각(delta content) 추출
    delta = choices[0].get("delta") or choices[0].get("message") or {}
    return delta.get("content")


def _stream_chat_completion(url, headers, data):
    # LLM API 엔드포인트로 HTTP POST 스트리밍 요청 전송
    with requests.post(url, headers=headers, json=data, stream=True, timeout=30) as response:
        # HTTP 에러 상태코드 체크
        response.raise_for_status()
        # 응답 인코딩 설정
        response.encoding = "utf-8"
        # 스트리밍 라인을 한 줄씩 수신
        for raw_line in response.iter_lines(decode_unicode=False):
            # 라인 파싱
            content = parse_stream_content(raw_line)
            # 스트림 종료 시 중단
            if content == "[DONE]":
                break
            # 텍스트 조각이 있으면 제너레이터로 yield 반환
            if content:
                yield content


def create_groq_response_stream(context, user_input):
    # Groq API 키가 없으면 예외 발생
    if not groq_api_key:
        raise ValueError("Groq API key is not configured.")

    # Groq OpenAI 호환 API 엔드포인트 URL
    url = "https://api.groq.com/openai/v1/chat/completions"
    # 인증 및 헤더 설정
    headers = {
        "Authorization": f"Bearer {groq_api_key}",
        "Content-Type": "application/json",
    }
    # 요청 파라미터 (LLaMA 3.3 70B 모델, 시스템 프롬프트, 낮은 temperature)
    data = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {"role": "system", "content": system_prompt.format(context=context)},
            {"role": "user", "content": user_input},
        ],
        "temperature": 0.1,
        "max_tokens": 1024,
        "stream": True,
    }
    # 스트리밍 함수 호출 반환
    return _stream_chat_completion(url, headers, data)


def create_openai_response_stream(context, user_input):
    # OpenAI API 키가 없으면 예외 발생
    if not openai_api_key:
        raise ValueError("OpenAI API key is not configured.")

    # OpenAI Chat Completions API 엔드포인트 URL
    url = "https://api.openai.com/v1/chat/completions"
    # 인증 및 헤더 설정
    headers = {
        "Authorization": f"Bearer {openai_api_key}",
        "Content-Type": "application/json",
    }
    # 요청 파라미터 (GPT-4o-mini 모델)
    data = {
        "model": "gpt-4o-mini",
        "messages": [
            {"role": "system", "content": system_prompt.format(context=context)},
            {"role": "user", "content": user_input},
        ],
        "temperature": 0.1,
        "max_tokens": 1024,
        "stream": True,
    }
    # 스트리밍 함수 호출 반환
    return _stream_chat_completion(url, headers, data)


def create_groq_response(context, user_input):
    # Groq 스트리밍 청크들을 하나로 병합하여 반환
    return "".join(create_groq_response_stream(context, user_input))


def create_openai_response(context, user_input):
    # OpenAI 스트리밍 청크들을 하나로 병합하여 반환
    return "".join(create_openai_response_stream(context, user_input))


# 규제 엑셀 데이터 로딩 스피너 UI 표시
with st.spinner("규제 지식베이스를 불러오는 중..."):
    # 엑셀 파일들로부터 문서 로드
    documents = load_documents()

# 로드된 문서들에 대한 의미론적 검색 백엔드(FAISS 인덱스 등) 구축
SEMANTIC_BACKEND = build_semantic_backend(documents)

# 지식베이스 문서가 하나도 없는 경우 경고 메시지 출력 후 앱 중단
if not documents:
    st.warning("현재 폴더에서 규제 문서를 찾지 못했습니다. 엑셀 파일이 있는지 확인해 주세요.")
    st.stop()


# LLM에 주입될 시스템 프롬프트 정의
system_prompt = (
    "당신은 M.I.Tech RA(Regulatory Affairs) 팀의 글로벌 의료기기 인허가 규제 전문 어시스턴트입니다.\n"
    "질문에 답할 때는 실무자처럼 자연스럽고 구체적으로 답하세요.\n"
    "아래 제공된 [참조 규정 정보]만을 엄격히 바탕으로 답변하고, 데이터에 없는 정보는 절대 지어내지 마세요.\n"
    "만약 참조 문서가 부족하면, 답변에 '확인된 정보만 기반으로 답변드립니다'라고 명시하고 불확실한 부분은 밝히세요.\n\n"
    "[참조 규정 정보]\n"
    "{context}\n\n"
    "답변 지침:\n"
    "1. [언어 통제] 반드시 100% 자연스러운 한국어로만 답변을 작성하세요 (영문 고유명사나 법령명은 제외). 'さらに' 같은 타 언어 접속사나 단어가 절대 섞이지 않도록 엄격히 주의하세요.\n"
    "2. [가독성 강화] 문장이 길어지거나 여러 세부 항목을 설명할 때는 반드시 줄바꿈(Enter)을 2번 이상 하여 문단을 확실히 띄우고, 글머리 기호(-, *, 1. 2.)를 적극적으로 사용하여 시각적으로 깔끔하게 분리하세요.\n"
    "3. [링크 안내] '자세한 내용은...' 이라며 URL 링크를 제공할 때는 앞 문장과 이어 쓰지 말고, 반드시 줄을 바꿔서 독립된 문단으로 출력하세요.\n"
    "4. 질문과 직접 관련된 문서만 요약하고, 각 국가/규제기관별로 항목을 나누어 설명하세요.\n"
    "5. 마지막에 '정리 및 주의사항' 섹션을 추가하고, '이는 참조 문서를 기반으로 한 분석이므로 정확한 해석은 담당자에게 검토를 받아주세요' 문구를 반드시 포함하세요.\n\n"
    "답변 스타일:\n"
    "• 국가/규제기관별로 명확하게 구분하세요\n"
    "• 핵심은 굵은 글씨(**)로 표시하세요\n"
    "• 비교 설명이 필요하면 두 국가의 차이점을 명확히 대조하여 작성하세요\n"
    "• 한눈에 이해할 수 있도록 구조화되고 시각적으로 읽기 편하게 작성하세요\n"
)

# LangChain ChatPromptTemplate 객체 생성
prompt = ChatPromptTemplate.from_messages([
    ("system", system_prompt),
    ("human", "{input}"),
])


# --- UI 및 메인 로직 영역 ---

# 세션 상태(Session State)에 대화 기록 리스트 초기화
if "messages" not in st.session_state:
    st.session_state.messages = []

# 세션 상태에 질문 답변 캐시 딕셔너리 초기화
if "answer_cache" not in st.session_state:
    st.session_state.answer_cache = {}

# 사이드바 국가 다중 선택 위젯
selected_countries = st.sidebar.multiselect("국가 필터", COUNTRY_OPTIONS[1:], default=[])
# 사이드바 변경 유형 다중 선택 위젯
selected_change_types = st.sidebar.multiselect("변경 유형 필터", CHANGE_TYPE_OPTIONS[1:], default=[])

# 사이드바 하단 설명 캡션
st.sidebar.caption("선택하지 않으면 전체 국가/전체 변경 유형으로 검색합니다.")

# 이전 대화 기록들을 화면에 차례대로 렌더링
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# 사용자 가이드용 예시 질문 출력
st.markdown("💡 **예시 질문:** 미국 FDA에서 라벨 변경 시 어떤 문서를 준비해야 하나요?")

# 사용자 채팅 입력창 처리
if user_input := st.chat_input("질문을 입력하세요"):
    # 1. 활성 필터 및 검색 질의어 재작성
    query_countries, query_change_type = get_active_filters(user_input, selected_countries, selected_change_types)
    rewritten_query = rewrite_query_for_search(user_input, selected_countries, selected_change_types)
    # 2. 질문 소스 분류(변경 규정 vs 일반 규정)
    query_source = get_query_source(rewritten_query)
    # 3. 최근 대화 메모리 구성
    conversation_memory = build_conversation_memory(st.session_state.messages)
    # 4. 미지원 국가 포함 여부 검사 (가드레일)
    blocked_country_question, blocked_country_names = should_block_country_question(user_input)

    # 사용자 질문을 세션 상태에 저장하고 화면에 렌더링
    st.session_state.messages.append({"role": "user", "content": user_input})
    with st.chat_message("user"):
        st.markdown(user_input)

    # 어시스턴트 답변 생성 영역 시작
    with st.chat_message("assistant"):
        with st.spinner("규제 DB 탐색 및 답변 생성 중..."):
            # 스트리밍 출력을 위한 빈 Streamlit 플레이스홀더 생성
            answer_placeholder = st.empty()
            selected_docs = []

            # 가드레일 조건 1: 미지원 국가 질문 차단
            if blocked_country_question:
                answer = build_country_scope_guard_response(blocked_country_names)
                answer_placeholder.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})
                st.session_state.last_api_key = "scope_guard"
            # 가드레일 조건 2: 의료기기 규제 도메인 외 질문 차단
            elif not is_global_regulatory_query(user_input):
                answer = build_scope_guard_response(user_input)
                answer_placeholder.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})
                st.session_state.last_api_key = "scope_guard"
            else:
                # 내부 지식베이스 문서 검색 실행 (상위 4건)
                selected_docs = retrieve_documents(
                    user_input,
                    documents,
                    top_k=4,
                    selected_country=selected_countries,
                    selected_change_type=selected_change_types,
                )

                # 외부 웹 검색 필요 여부 판단
                query_clean = normalize_text(user_input)
                web_triggers = ["최신", "뉴스", "동향", "가이드라인", "업데이트", "발표", "현재"]
                # 엑셀에 문서가 없거나, '최신' 등의 단어가 포함되었을 때만 Tavily 가동
                needs_web_search = not selected_docs or any(t in query_clean for t in web_triggers)
                
                external_context = ""
                # 외부 웹 검색 실행
                if needs_web_search:
                    with st.spinner("내부 규정에 없는 최신 규제/동향 정보를 웹에서 검색 중입니다..."):
                        external_context = search_external_regulatory_info(user_input)

                # 캐시 키 생성 및 기존 캐시 조회
                cache_key = build_cache_key(user_input, selected_countries, selected_change_types, selected_docs, is_simple_question(user_input))
                cached_answer = get_cached_answer(st.session_state.answer_cache, cache_key)

                try:
                    # 분기 1: 캐시된 답변이 있는 경우 즉시 출력
                    if cached_answer:
                        answer = cached_answer
                        st.session_state.last_api_key = "cache"
                        answer_placeholder.markdown(answer)
                    # 분기 2: 단순 조회 질문이며 외부 검색 결과가 없는 경우 로컬 파일 기반으로 즉시 요약 답변 생성
                    elif is_simple_question(user_input) and selected_docs and not external_context:
                        answer = (
                            "단순 정보 조회로 확인되어 자체 데이터베이스 기반으로 답변드립니다.\n\n"
                            + build_file_based_answer(user_input, selected_docs)
                        )
                        answer = finalize_answer_text(answer, selected_docs)
                        answer_placeholder.markdown(answer)
                        st.session_state.last_api_key = "local"
                    else:
                        # 분기 3: 내부 문서와 외부 검색 결과가 모두 없는 경우 처리
                        if not selected_docs and not external_context:
                            answer = "해당 질문과 관련된 규정 문서를 내부 DB 및 웹 검색에서 찾을 수 없습니다."
                            answer_placeholder.markdown(answer)
                            st.session_state.last_api_key = "local"
                        else:
                            # 분기 4: LLM API 호출을 위한 종합 컨텍스트 생성
                            context = build_retrieval_context(
                                user_input,
                                selected_docs,
                                conversation_memory=conversation_memory,
                                rewritten_query=rewritten_query,
                            )

                            # 외부 웹 검색 결과가 있으면 컨텍스트에 추가 병합
                            if external_context:
                                context += f"\n\n{external_context}"
                                
                            answer = None
                            last_error = None
                            used_key = None

                            # Groq 우선 호출 -> 실패 시 OpenAI로 순차적 폴백
                            for provider_name, stream_fn in [
                                ("groq", create_groq_response_stream),
                                ("openai", create_openai_response_stream),
                            ]:
                                try:
                                    streamed_text = ""
                                    # 스트리밍 텍스트 조각을 플레이스홀더에 실시간 업데이트
                                    for chunk in stream_fn(context, user_input):
                                        streamed_text += chunk
                                        answer_placeholder.markdown(streamed_text)
                                    # 최종 완성된 답변에 원문/번역 및 URL 정제 적용
                                    answer = finalize_answer_text(streamed_text, selected_docs)
                                    answer_placeholder.markdown(answer)
                                    used_key = provider_name
                                    break
                                except Exception as provider_error:
                                    last_error = provider_error

                            # 모든 LLM 공급자 호출이 실패했을 경우 예외 발생
                            if not used_key:
                                raise last_error

                            # 성공한 공급자 이름 기록
                            st.session_state.last_api_key = used_key

                        # 생성된 답변을 캐시에 저장
                        set_cached_answer(st.session_state.answer_cache, cache_key, answer)
                    # 어시스턴트 메시지를 대화 기록에 저장
                    st.session_state.messages.append({"role": "assistant", "content": answer})
                except Exception as e:
                    # 모든 API 호출 실패 시 로컬 문서 요약으로 폴백 처리
                    fallback_answer = (
                        "⚠️ 현재 외부 AI 서버 연결이 원활하지 않아 상세 비교 분석을 제공할 수 없습니다. "
                        "대신 관련된 원문 요약 정보를 제공해 드립니다.\n\n"
                        + build_file_based_answer(user_input, selected_docs)
                    )
                    fallback_answer = finalize_answer_text(fallback_answer, selected_docs)
                    answer_placeholder.markdown(fallback_answer)
                    st.session_state.messages.append({"role": "assistant", "content": fallback_answer})
                    # 에러 메시지 출력
                    st.error(f"내부 오류 발생: 모든 AI 모델 호출 실패 ({e})")

        # 인수인계자 및 개발자를 위한 실시간 디버그 정보 아코디언 영역
        with st.expander("🔍 디버그 정보"):
            st.write(f"**질문 분류:** {'단순 조회 (자체 요약)' if is_simple_question(user_input) else '복합 분석 (Groq/OpenAI API 호출)'}")
            st.write(f"**검색 출처:** {query_source}")
            st.write(f"**재작성 질문:** {rewritten_query}")
            st.write(f"**감지된 국가:** {', '.join(query_countries) if query_countries else '없음'}")
            st.write(f"**감지된 변경 유형:** {', '.join(query_change_type) if query_change_type else '없음'}")
            st.write(f"**검색된 문서 수:** {len(selected_docs)}")
            if "last_api_key" in st.session_state:
                st.write(f"**활성화된 AI 제공자:** {str(st.session_state.last_api_key).upper()}")
            if selected_docs:
                st.write("**참고한 규제 원문:**")
                for i, doc in enumerate(selected_docs, 1):
                    preview = doc['content'][:150].replace('\n', ' ')
                    st.write(f"{i}. {preview}...")