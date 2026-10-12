# 네이버 금융 분류코드(etfTabCode)별 의미

# 분류코드 숫자 / 네이버 금융 매칭 탭 이름 / 해당 카테고리 설명 / 대표 종목 예시
# 1	/ 국내 시장대표 / 코스피200, 코스닥150 등 국내 대표 주가지수를 그대로 추종하는 상품	/ KODEX 200, TIGER 코스닥150
# 2	/ 국내 테마 / 반도체, AI, 2차전지, 바이오 등 국내 특정 산업군에 집중 투자하는 상품	/ TIGER Fn반도체, KODEX 200중소형
# 3	/ 해외 주식 / 미국, 일본, 유럽 등 해외 국가의 지수나 테마에 투자하는 상품	/ TIGER 미국S&P500, ACE 미국빅테크10
# 4	/ 단기자금 / CD금리, KOFR금리 등 파킹통장처럼 안정적인 이자 수익을 추구하는 상품	/ KODEX CD금리액티브, TIGER KOFR금리액티브
# 5	/ 채권 / 국내외 국채, 회사채 등 채권 자산에 투자하여 변동성을 낮춘 상품	/ KODEX 국고채3년, TIGER 미국채10년
# 6	/ 기타 / 원자재(금, 은, 원유), 리츠, 통화(달러 인버스/레버리지) 등 대안 자산 상품	/ KODEX 골드선물(H), TIGER 미국달러선물


import pandas as pd
import requests
import json
from datetime import datetime, timedelta
import urllib3
import FinanceDataReader as fdr

# SSL 보안 무시 경고 메시지 비활성화
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_asset_manager(etf_name):
    """ETF 명칭을 분석하여 자산운용사를 매칭합니다."""
    brands = {
        'KODEX': '삼성자산운용', 'TIGER': '미래에셋자산운용', 'KBSTAR': 'KB자산운용',
        'ACE': '한국투자신탁운용', 'SOL': '신한자산운용', 'PLUS': '한화자산운용',
        'HANARO': '우리자산운용', 'KOSEF': '키움투자자산운용', 'TIMEFOLIO': '타임폴리오자산운용'
    }
    for brand, manager in brands.items():
        if brand in str(etf_name).upper():
            return manager
    return '기타/외국계'

def save_all_etf_fields_to_excel():
    today = datetime.today().strftime("%Y%m%d_%H%M")
    
    # 2. 정확히 30일 전(1개월 전) 날짜 산출 (주가는 영업일 기준으로 자동 조회됨)
    start_date = (datetime.today() - timedelta(days=90)).strftime('%Y-%m-%d')
    
    print(f"[{today}] 국내 상장된 모든 ETF의 전방위 금융 지표 추출 작업을 시작합니다...")
    
    url = "https://finance.naver.com/api/sise/etfItemList.nhn"
    
    # 네이버 웹 차단벽을 무력화하기 위한 실제 브라우저 풀 패키지 요청 헤더
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'Referer': 'https://naver.com',
        'Origin': 'https://finance.naver.com'
    }
    
    # 전달해주신 분류 기준에 맞춘 네이버 금융 한글 매칭 탭 딕셔너리 구성
    tab_mapping = {
        1: '국내 시장대표',
        2: '국내 테마',
        3: '해외 주식',
        4: '단기자금',
        5: '채권',
        6: '기타'
    }
    
    try:
        # 1. 브라우저 세션 상태를 유지하여 데이터 수집 통로 안정화
        session = requests.Session()
        response = session.get(url, headers=headers, verify=False, timeout=15)
        
        # 2. 한글 디코딩 및 공백 문자 에러 제어
        response.encoding = 'cp949'
        response_text = response.text.strip()
        
        # 만약 차단 HTML 페이지가 반환되었을 경우를 대비한 2차 우회 로직
        if response_text.startswith('<') or not response_text.startswith('{'):
            alt_url = "https://naver.com"
            response = session.get(alt_url, headers=headers, verify=False, timeout=15)
            response.encoding = 'cp949'
            response_text = response.text.strip()
            
        # 3. JSON 원천 데이터셋 해석
        json_data = json.loads(response_text)
        etf_list = json_data['result']['etfItemList']
        etf_list_len = len(etf_list)
        
        all_indicators = []
        len_count = 0
        
        # 4. API에 내장된 모든 유효 데이터 매핑 전처리
        for item in etf_list:
            name = item.get('itemname', '')
            tab_code = item.get('etfTabCode', 0) # 원본 API에서 분류코드 추출
            
            all_indicators.append({
                "종목코드": str(item.get('itemcode', '')).zfill(6),
                "ETF명": name,
                "자산운용사": get_asset_manager(name),
                "순자산가치(NAV)": item.get('nav', 0),
                "현재가(원)": item.get('nowVal', 0),
                "전일비(원)": item.get('changeVal', 0),
                "등락률(%)": item.get('changeRate', 0.0),
                "3개월수익률(%)": item.get('threeMonthEarnRate', 0.0),
                "2개월수익률(%)": 0.0,
                "1개월수익률(%)": 0.0,
                "3주수익률(%)": 0.0,
                "2주수익률(%)": 0.0,
                "1주수익률(%)": 0.0,
                "거래량(주)": item.get('quant', 0),
                "거래대금(원)": item.get('amonut', 0),  # 네이버 내부 오타 키값인 amonut 준수
                "시가총액(억)": item.get('marketSum', 0),
                "분류코드(Tab)": tab_code,
                "분류명(네이버탭)": tab_mapping.get(int(tab_code), '미분류') # 🌟 원본 변형 없이 한글 매칭 열만 추가 장착
            })
            
            print(f"{len_count}/{etf_list_len}.....{name}")
            
            try:
                # 개별 종목의 최근 한 달간의 종가(Close) 가져오기
                price_df = fdr.DataReader(item.get('itemcode', ''), start_date)
                previous_price = price_df['Close'].iloc[0]   # 한 달 전 종가
                today_price = price_df['Close'].iloc[-1] # 현재(오늘) 종가
                
                if len(price_df) >= 2:
                    return_rate = round(((today_price - previous_price) / previous_price) * 100, 2)                    
                    all_indicators[-1]["2개월수익률(%)"] = return_rate
                    
                if len(price_df) >= 7 :
                    # 1개월 변동 수익률 연산
                    previous_price   = price_df['Close'].iloc[-7]   
                    return_rate = round(((today_price - previous_price) / previous_price) * 100, 2)                    
                    all_indicators[-1]["1주수익률(%)"] = return_rate

                if len(price_df) >= 14 :
                    # 1개월 변동 수익률 연산
                    previous_price   = price_df['Close'].iloc[-14]   
                    return_rate = round(((today_price - previous_price) / previous_price) * 100, 2)                    
                    all_indicators[-1]["2주수익률(%)"] = return_rate

                if len(price_df) >= 21 :
                    # 1개월 변동 수익률 연산
                    previous_price   = price_df['Close'].iloc[-21]   
                    return_rate = round(((today_price - previous_price) / previous_price) * 100, 2)                    
                    all_indicators[-1]["3주수익률(%)"] = return_rate

                if len(price_df) >= 30 :
                    # 1개월 변동 수익률 연산
                    previous_price   = price_df['Close'].iloc[-30]   
                    return_rate = round(((today_price - previous_price) / previous_price) * 100, 2)                    
                    all_indicators[-1]["1개월수익률(%)"] = return_rate

                if len(price_df) >= 60 :
                    # 2개월 변동 수익률 연산
                    previous_price = price_df['Close'].iloc[-60]   
                    return_rate = round(((today_price - previous_price) / previous_price) * 100, 2)                    
                    all_indicators[-1]["2개월수익률(%)"] = return_rate

            except Exception as error :
                print(f"error .... {error}")
                pass # 상장폐지 혹은 데이터 지연 종목 패스            

            len_count = len_count + 1 
            
            # if len_count > 10 :
            #     break
            
        # 5. 판다스 변환 및 데이터 가독성 정렬 (기본값: 거래대금 많은 순)
        df_final = pd.DataFrame(all_indicators)
        df_final = df_final.sort_values(by='거래대금(원)', ascending=False).reset_index(drop=True)
        
        # 6. 테마 필터링용 데이터프레임 구성
        df_semiconductor = df_final[df_final['ETF명'].str.contains('반도체|네오셈|한미반도체', case=False, na=False)]
        df_ai = df_final[df_final['ETF명'].str.contains('AI|인공지능|테크|빅테크', case=False, na=False)]
        df_dividend = df_final[df_final['ETF명'].str.contains('배당|고배당|인컴', case=False, na=False)]
        df_bond = df_final[df_final['ETF명'].str.contains('채권|국채|회사채|통화', case=False, na=False)]
        
        # 7. 단일 통합 엑셀 마스터 파일 내 다중 시트 저장
        file_name = f"naver_etf_all_indicators_{today}.xlsx"
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='전체 ETF 지표', index=False)
            df_semiconductor.to_excel(writer, sheet_name='반도체 테마', index=False)
            df_ai.to_excel(writer, sheet_name='AI & 빅테크', index=False)
            df_dividend.to_excel(writer, sheet_name='배당 & 고배당', index=False)
            df_bond.to_excel(writer, sheet_name='채권형', index=False)
            
        print("=" * 70)
        print(f"🎉 [성공] 모든 ETF의 마스터 금융 지표가 안전하게 통합 보존되었습니다.")
        print(f"📁 생성 파일명: {file_name}")
        print(f"📊 수집 대상 총계: 국내 상장 ETF 총 {len(df_final)}개 종목의 전체 컬럼 추출 완료")
        print("=" * 70)
        
    except Exception as e:
        print(f"❌ 데이터 최종 변환 및 저장 과정 중 오류 발생: {e}")

if __name__ == "__main__":
    save_all_etf_fields_to_excel()
