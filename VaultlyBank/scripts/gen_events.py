# Generates Vaultly/Models/EventLibrary.swift (run from VaultlyBank/). Amounts are middle-class KRW; tier scale applied in-app.
# Counterparties use {bank}/{broker}/{card}/{capital} placeholders — resolved from user-editable 금융사 이름 at runtime.
# kind: ("d", chance) daily-random, ("m", day) monthly, ("y", month, day) yearly
# tiers: subset of B W M R T (거지 서민 중산층 부자 재벌). cat: transfer/interest/repayment/autoTransfer
ALL="BWMRT"
S=[]
# scale: "full" = multiply by tierScale (income, taxes, assets); "fixed" = same price for every tier
#        (GS25, taxi, pharmacy — or bases already written at that tier's level); number = cap on tierScale
#        (a rich person eats at nicer places, but a 맥도날드 bill is never 1,000,000원)
def spec(id, kind, tiers, lo, hi, cps, memos, cat="transfer", hour=12, scale="full"):
    S.append(dict(id=id,kind=kind,tiers=tiers,lo=lo,hi=hi,cps=cps,memos=memos,cat=cat,hour=hour,scale=scale))

# ---------- 보험 (monthly) ----------
spec("ins_health",("m",5),ALL,-98000,-145000,["삼성화재","현대해상","DB손해보험","KB손해보험","메리츠화재"],["실손보험료","실비보험 자동이체","건강보험 월납","실손 갱신보험료"],"autoTransfer",8,scale=2)
spec("ins_life",("m",10),"WMRT",-120000,-380000,["삼성생명","한화생명","교보생명","신한라이프","NH농협생명"],["종신보험료","생명보험 월납","변액보험 납입","연금보험 납입"],"autoTransfer",8,scale=3)
spec("ins_car",("m",15),"MRT",-65000,-140000,["DB손해보험 다이렉트","삼성화재 다이렉트","현대해상 하이카","AXA손해보험"],["자동차보험 분납","자동차보험료","차량보험 월납"],"autoTransfer",8,scale=2.5)
spec("ins_fire",("m",20),"MRT",-12000,-48000,["KB손해보험","메리츠화재","롯데손해보험"],["화재보험료","주택화재보험","가재도구 보험"],"autoTransfer",8,scale=2)
spec("ins_dental",("m",17),ALL,-19000,-42000,["라이나생명","AIA생명","처브라이프"],["치아보험료","치과보험 월납"],"autoTransfer",8,scale="fixed")
spec("ins_cancer",("m",22),"WMRT",-35000,-89000,["라이나생명","동양생명","흥국생명","ABL생명"],["암보험료","3대질병 보험료","뇌심혈관 보험료"],"autoTransfer",8,scale=2)
spec("ins_pet",("m",12),"MRT",-29000,-68000,["펫퍼민트","메리츠 펫보험","DB 펫블리"],["반려동물 보험료","펫보험 월납"],"autoTransfer",8,scale="fixed")
spec("ins_payout",("d",0.004),ALL,180000,2400000,["삼성화재","현대해상","DB손해보험","KB손해보험","메리츠화재","한화손해보험"],["실손보험금 지급","보험금 입금","진단비 지급","입원일당 지급","통원치료비 지급"],"transfer",14,scale=3)
spec("ins_refund",("d",0.002),"WMRT",40000,520000,["삼성생명","한화생명","교보생명","신한라이프"],["보험 환급금","만기 환급금","배당금 지급","해지환급금"],"transfer",10,scale=3)

# ---------- 공과금/구독 (monthly) ----------
spec("util_elec",("m",25),ALL,-28000,-92000,["한국전력공사"],["전기요금","전기요금 자동납부","전기세"],"autoTransfer",9,scale=3)
spec("util_gas",("m",23),ALL,-12000,-110000,["서울도시가스","삼천리","예스코","경동도시가스","코원에너지"],["도시가스 요금","가스요금 자동이체","가스비"],"autoTransfer",9,scale=3)
spec("util_water",("m",18),"WMRT",-9000,-34000,["서울아리수본부","수도사업소","K-water"],["수도요금","상하수도 요금"],"autoTransfer",9,scale=2)
spec("util_net",("m",8),ALL,-22000,-49000,["KT", "SK브로드밴드", "LG유플러스", "딜라이브"],["인터넷 요금","인터넷+TV 결합","IPTV 요금"],"autoTransfer",9,scale="fixed")
spec("util_phone",("m",20),ALL,-35000,-110000,["SK텔레콤", "KT", "LG유플러스", "KT M모바일", "토스모바일"],["휴대폰 요금","통신요금 자동납부","휴대폰 할부+요금"],"autoTransfer",7,scale=1.5)
spec("sub_ott",("m",3),ALL,-5500,-17000,["넷플릭스", "디즈니+", "쿠팡 와우멤버십", "티빙", "웨이브", "유튜브 프리미엄", "왓챠"],["월 구독료","정기결제","멤버십 결제"],"autoTransfer",6,scale="fixed")
spec("sub_music",("m",7),ALL,-6900,-10900,["멜론","스포티파이","지니뮤직","애플뮤직"],["월 구독료","정기결제"],"autoTransfer",6,scale="fixed")
spec("sub_cloud",("m",11),"WMRT",-1100,-14000,["iCloud","Google One","네이버 MYBOX","Dropbox"],["저장공간 구독","클라우드 요금"],"autoTransfer",6,scale="fixed")
spec("sub_gym",("m",1),"WMRT",-49000,-180000,["스포애니 서초점", "에이블짐 강남점", "커브스 반포점", "F45 역삼", "바디채널 사당점"],["헬스장 회비","PT 분납","필라테스 월회비","수영장 월회비"],"autoTransfer",10,scale="fixed")
spec("sub_app",("d",0.03),ALL,-1200,-14900,["Apple App Store", "Google Play", "Steam", "닌텐도 eShop"],["앱 결제","인앱 구매","게임 아이템","유료 앱 구매","배틀패스"],"transfer",22,scale="fixed")
spec("mgmt_fee",("m",10),"MRT",-180000,-620000,["래미안퍼스티지 관리사무소", "아크로리버파크 관리사무소", "자이 관리사무소", "힐스테이트 관리사무소"],["관리비","아파트 관리비","관리비+주차비","장기수선충당금 포함 관리비"],"autoTransfer",9,scale=4)

# ---------- 세금 (yearly/monthly) ----------
spec("tax_property1",("y",7,25),"MRT",-280000,-1900000,["위택스","서울시 세무과","강남구청","송파구청","용산구청"],["재산세 (주택 1기분)","재산세 7월분","재산세 건물분"],"transfer",10)
spec("tax_property2",("y",9,25),"MRT",-280000,-1900000,["위택스","서울시 세무과","강남구청","송파구청","용산구청"],["재산세 (주택 2기분)","재산세 9월분","재산세 토지분"],"transfer",10)
spec("tax_car",("y",6,20),"MRT",-180000,-520000,["위택스","차량등록사업소"],["자동차세 1기분","자동차세 연납"],"transfer",10,scale=3)
spec("tax_car2",("y",12,20),"MRT",-180000,-520000,["위택스","차량등록사업소"],["자동차세 2기분"],"transfer",10,scale=3)
spec("tax_resident",("y",8,25),ALL,-6000,-12500,["위택스","구청 세무과"],["주민세","개인분 주민세"],"transfer",10,scale="fixed")
spec("tax_refund",("y",3,12),"WMRT",150000,1800000,["국세청","홈택스"],["연말정산 환급","근로소득 환급금","종합소득세 환급"],"transfer",9,scale=3)
spec("tax_yearend_pay",("y",2,28),"MRT",-200000,-1500000,["국세청","홈택스"],["연말정산 추가납부","소득세 추가 징수"],"transfer",9)
spec("tax_compre",("y",5,31),"RT",-4800000,-32000000,["국세청","홈택스"],["종합소득세 납부","종합소득세 확정신고 납부"],"transfer",10,scale="fixed")
spec("tax_cpt",("y",12,15),"RT",-2800000,-22000000,["국세청"],["종합부동산세","종부세 납부"],"transfer",10,scale="fixed")
spec("pension",("m",10),"BW",-90000,-210000,["국민연금공단"],["국민연금 지역가입 보험료","국민연금 납부"],"autoTransfer",9,scale="fixed")
spec("nhi",("m",10),"BW",-45000,-130000,["국민건강보험공단"],["건강보험료 (지역가입)","건보료 납부","장기요양보험료 포함 건보료"],"autoTransfer",9,scale="fixed")

# ---------- 의료 ----------
spec("med_clinic",("d",0.045),ALL,-4500,-38000,["연세내과","서울이비인후과","하나정형외과","참좋은피부과","우리가정의학과","밝은안과"],["진료비","외래 진료","물리치료","감기 진료","재진"],"transfer",11,scale="fixed")
spec("med_pharm",("d",0.05),ALL,-3200,-24000,["온누리약국 서초점", "메디팜약국", "굿모닝약국", "참약국", "건강약국"],["처방약","약값","감기약","비타민","영양제"],"transfer",12,scale="fixed")
spec("med_dental",("d",0.012),ALL,-15000,-480000,["서울밝은치과","연세플란트치과","미소드림치과","강남예치과"],["스케일링","충치 치료","신경치료","크라운","레진","사랑니 발치"],"transfer",14,scale=1.5)
spec("med_er",("d",0.003),ALL,-85000,-640000,["서울대병원 응급실","세브란스 응급의료센터","아산병원 응급실","삼성서울병원"],["응급실 진료비","야간 응급 진료","CT 촬영 포함 응급진료"],"transfer",2,scale="fixed")
spec("med_surgery",("d",0.0012),ALL,-800000,-6500000,["서울아산병원","삼성서울병원","세브란스병원","서울대병원","강남세브란스"],["입원 치료비","수술비","입원비 정산","검사+입원비"],"transfer",10,scale=2.5)
spec("med_checkup",("y",10,14),"WMRT",-90000,-1200000,["KMI한국의학연구소","하나로의료재단","강북삼성병원 건진센터","차움"],["종합검진","건강검진 추가검사","프리미엄 건진"],"transfer",9,scale=3)
spec("med_vet",("d",0.015),"MRT",-25000,-420000,["24시 동물병원","이리온 동물병원","해마루 동물병원"],["반려견 진료","고양이 예방접종","반려동물 치료비","중성화 수술","스케일링(반려견)"],"transfer",15,scale=1.5)
spec("med_mental",("m",14),"WMRT",-60000,-180000,["마음숲 정신건강의학과","연세 심리상담센터"],["상담 치료","심리상담","정신과 진료"],"transfer",17,scale="fixed")

# ---------- 사고/과태료/벌금 ----------
spec("fine_traffic",("d",0.008),"WMRT",-32000,-130000,["경찰청 교통민원24","서울경찰청","도로교통공단"],["속도위반 과태료","신호위반 범칙금","주정차 위반 과태료","버스전용차로 위반","안전벨트 미착용"],"transfer",10,scale="fixed")
spec("fine_parking",("d",0.01),"MRT",-40000,-80000,["강남구청","송파구청","서초구청","마포구청"],["주정차 위반 과태료","불법주차 과태료"],"transfer",10,scale="fixed")
spec("accident_pay",("d",0.0025),"WMRT",-300000,-4800000,["상대 차량 운전자","보험사 구상금","정비공장"],["교통사고 합의금","자차 수리비 자기부담금","접촉사고 수리비","사고 처리비"],"transfer",16,scale=3)
spec("accident_recv",("d",0.0015),"WMRT",250000,3200000,["삼성화재 대인보상","DB손보 보상팀","현대해상 보상센터"],["교통사고 합의금 수령","대인 보상금","대물 보상금","휴업손해 배상"],"transfer",14,scale=3)
spec("fine_late",("d",0.02),"B",-5000,-45000,["{capital} 채권관리팀", "(주)러시앤캐시 추심", "OO신용정보"],["연체료","독촉 수수료","연체이자 추가","추심 비용"],"interest",8,scale="fixed")
spec("fine_misc",("d",0.004),ALL,-20000,-200000,["법원","지방법원 공탁","경찰서"],["과태료","벌금","즉결심판 벌금","공탁금"],"transfer",10,scale=3)
spec("car_repair",("d",0.006),"MRT",-80000,-1800000,["현대 블루핸즈 강남점", "기아 오토큐 서초점", "한성자동차 벤츠 서비스센터", "코오롱모터스 BMW 서비스", "스피드메이트 방배점", "타이어뱅크 양재점"],["엔진오일 교체","타이어 교체","브레이크 패드","정기 점검","범퍼 수리","배터리 교체","에어컨 수리"],"transfer",15,scale=3)
spec("car_fuel",("d",0.12),"MRT",-45000,-120000,["SK에너지 반포주유소", "GS칼텍스 양재주유소", "현대오일뱅크 신사주유소", "S-OIL 송파셀프주유소", "테슬라 슈퍼차저", "환경부 전기차 충전"],["주유","경유 주유","전기차 충전","세차+주유"],"transfer",8,scale=1.5)
spec("car_toll",("d",0.06),"MRT",-2000,-24000,["한국도로공사", "하이패스"],["통행료","고속도로 통행료","하이패스 충전"],"transfer",9,scale="fixed")

# ---------- 경조사 ----------
spec("wedding",("d",0.006),ALL,-50000,-300000,["김OO 결혼식","박OO 결혼식","동기 결혼","사촌 결혼식","회사 후배 결혼"],["축의금","결혼 축의금","축의금 송금"],"transfer",12,scale=3)
spec("funeral",("d",0.003),ALL,-50000,-300000,["장례식장","이OO 부친상","최OO 모친상","친구 조부상"],["부의금","조의금","부의금 송금"],"transfer",19,scale=3)
spec("baby",("d",0.003),ALL,-30000,-200000,["친구 돌잔치","조카 돌","동료 출산"],["돌잔치 축하금","출산 축하금","돌반지 구입"],"transfer",13,scale=3)
spec("gift_recv",("d",0.004),ALL,30000,300000,["엄마","아빠","이모","삼촌","할머니","고모","외삼촌"],["생일 축하금","용돈","설 용돈","추석 용돈","격려금"],"transfer",10,scale=5)
spec("gift_send",("d",0.004),"WMRT",-50000,-500000,["엄마","아빠","장모님","장인어른","시어머니","시아버지"],["용돈 드림","생신 축하","어버이날","명절 용돈","효도 송금"],"transfer",11,scale=5)
spec("holiday_bonus",("y",9,20),"WMRT",200000,1500000,["(주)회사","인사팀"],["추석 상여금","명절 상여","명절 선물비"],"transfer",9)
spec("holiday_bonus2",("y",1,25),"WMRT",200000,1500000,["(주)회사","인사팀"],["설 상여금","설 명절 상여"],"transfer",9)

# ---------- 급여 외 소득 ----------
spec("bonus_perf",("y",3,20),"WMRT",800000,9000000,["(주)회사","인사팀","재무팀"],["성과급","PS 지급","인센티브","연간 성과 보너스"],"transfer",9)
spec("bonus_semi",("y",7,20),"MRT",500000,4000000,["(주)회사","인사팀"],["상반기 인센티브","PI 지급","하계 상여"],"transfer",9)
spec("overtime",("m",10),"BWM",80000,450000,["(주)회사 급여팀","급여팀"],["초과근무수당","야근수당","휴일근로수당","연장근로 수당"],"transfer",9,scale="fixed")
spec("side_income",("d",0.02),"BWM",15000,320000,["쿠팡플렉스","배민커넥트","당근마켓","번개장터","크몽","숨고","알바천국"],["부업 정산","중고 판매","배달 수익","프리랜서 대금","단기 알바비","원고료"],"transfer",18,scale="fixed")
spec("refund",("d",0.025),ALL,8000,180000,["쿠팡(주)", "(주)무신사", "네이버페이", "G마켓", "11번가", "올리브영"],["환불","주문 취소 환불","반품 환불","부분 환불"],"transfer",13,scale=3)
spec("cashback",("d",0.04),ALL,500,25000,["토스", "네이버페이", "카카오페이", "페이코", "{card}"],["캐시백","포인트 현금화","이벤트 리워드","결제 캐시백","만보기 리워드"],"transfer",21,scale="fixed")
spec("lottery",("d",0.006),ALL,5000,50000,["동행복권","로또","스피또"],["로또 5등 당첨","즉석복권 당첨","복권 4등 당첨"],"transfer",20,scale="fixed")
spec("lottery_big",("d",0.0003),ALL,1500000,30000000,["동행복권","연금복권"],["로또 3등 당첨금","연금복권 2등","로또 2등 당첨금"],"transfer",20,scale="fixed")
spec("dividend",("d",0.03),"MRT",12000,1800000,["삼성전자","SK하이닉스","KB금융","현대차","맥쿼리인프라","리얼티인컴","SCHD","코카콜라"],["배당금 입금","분기 배당","월 배당 ETF","결산 배당"],"interest",15)
spec("interest_in",("m",1),"WMRT",2400,480000,["{bank}", "카카오뱅크", "토스뱅크", "케이뱅크"],["예금 이자","파킹통장 이자","CMA 이자","적금 이자"],"interest",0)
spec("rent_in",("m",5),"RT",1200000,9500000,["세입자 김OO","세입자 박OO","상가 임차인","오피스텔 세입자","원룸 세입자"],["월세 입금","임대료","상가 임대료","오피스텔 월세"],"transfer",10,scale="fixed")
spec("rent_deposit_in",("d",0.001),"RT",50000000,500000000,["신규 세입자","부동산 중개"],["전세 보증금 입금","임대 보증금 수령"],"transfer",14,scale="fixed")
spec("rent_deposit_out",("d",0.001),"RT",-50000000,-500000000,["퇴거 세입자","부동산 중개"],["전세 보증금 반환","임대 보증금 반환"],"transfer",14,scale="fixed")
spec("biz_profit",("d",0.08),"T",8000000,450000000,["VX벤처스","VX캐피탈","(주)VX물산","VX바이오","VX에너지","해외 펀드 LP","PEF 운용사"],["투자 수익 배분","펀드 분배금","계열사 배당","지분 매각 대금","성과 보수","엑시트 수익","우선주 배당"],"transfer",11,scale="fixed")
spec("biz_profit_r",("d",0.02),"R",2000000,45000000,["스타트업 투자조합","부동산 펀드","P2P 플랫폼","개인투자조합"],["투자 수익 배분","펀드 분배금","엑시트 수익","지분 매각"],"transfer",11,scale="fixed")
spec("royalty",("m",15),"RT",800000,35000000,["음원 유통사","출판사","특허 라이선시","브랜드 라이선스"],["저작권료","인세","특허 로열티","브랜드 사용료"],"transfer",10,scale="fixed")
spec("consult",("d",0.03),"RT",3000000,40000000,["(주)컨설팅 고객사","사외이사 보수","강연 주최측","자문사"],["자문료","사외이사 보수","강연료","고문료"],"transfer",14,scale="fixed")

# ---------- 대출/채무 ----------
spec("loan_student",("m",15),"BW",-120000,-280000,["한국장학재단"],["학자금 대출 상환","학자금 원리금"],"repayment",9,scale="fixed")
spec("loan_card",("m",9),"BW",-150000,-600000,["{card} 카드론", "현대카드 론", "롯데카드"],["카드론 상환","카드 리볼빙","현금서비스 상환","카드론 이자"],"repayment",8,scale="fixed")
spec("loan_capital",("m",6),"B",-180000,-450000,["{capital}", "JB우리캐피탈", "하나캐피탈", "리드코프", "러시앤캐시"],["대출 원금 상환","원리금 상환","자동차 담보대출 상환"],"repayment",8,scale="fixed")
spec("loan_private",("d",0.03),"B",-30000,-200000,["사채업자","일수","지인 채권자","OO대부"],["일수 상환","사채 이자","빌린 돈 이자","독촉 상환"],"interest",7,scale="fixed")
spec("loan_mortgage",("m",10),"MR",-900000,-3800000,["KB국민은행","신한은행","우리은행","하나은행"],["주택담보대출 원리금","주담대 상환","전세자금대출 이자"],"repayment",9)
spec("loan_recv",("d",0.004),"BW",100000,2000000,["지인","친구","형","누나","부모님"],["돈 빌림","급전 빌림","생활비 지원","빌려준 돈 돌려받음"],"transfer",21,scale=3)
spec("loan_give",("d",0.004),"MRT",-200000,-5000000,["친구","동생","사촌","후배"],["돈 빌려줌","급전 빌려줌","지인 대여"],"transfer",21,scale=3)

# ---------- 생활 지출 ----------
spec("grocery",("d",0.35),ALL,-8000,-120000,["이마트 성수점", "홈플러스 강서점", "롯데마트 서울역점", "쿠팡(주)", "(주)컬리", "코스트코 양재점", "GS더프레시 방배점", "노브랜드 신촌점", "이마트 에브리데이 반포점"],["장보기","식재료","생필품","주간 장보기","로켓배송"],"transfer",17,scale=2)
spec("dining",("d",0.6),ALL,-6000,-85000,["맘스터치 신논현점", "김밥천국 서초점", "스타벅스 강남R점", "(주)우아한형제들", "쿠팡이츠", "요기요", "(주)맥도날드 삼성중앙역점", "서브웨이 테헤란로점", "이디야커피 신사역점", "버거킹 강남대로점", "롯데리아 사당역점", "본죽 역삼점", "한솥도시락 선릉점", "교촌치킨 반포점", "투썸플레이스 종로점"],["식사","점심","저녁","배달","커피","야식","브런치"],"transfer",12,scale=1.5)
spec("dining_fine",("d",0.08),"RT",-180000,-2400000,["정식당","밍글스","스시 효","라연","모수","오마카세 스시조"],["디너 코스","오마카세","와인 페어링","비즈니스 디너","접대"],"transfer",20,scale="fixed")
spec("convenience",("d",0.5),ALL,-1500,-18000,["GS25 서초점", "CU 역삼센트럴점", "세븐일레븐 논현역점", "이마트24 성수점", "GS25 마포공덕점", "CU 잠실새내점", "세븐일레븐 신촌점"],["편의점","담배","음료","간식","도시락"],"transfer",22,scale="fixed")
spec("transport",("d",0.4),"BWM",-1400,-15000,["티머니", "서울교통공사", "카카오T 택시", "우버 택시", "타다", "서울시 따릉이"],["대중교통","지하철","버스","택시","공유자전거"],"transfer",8,scale="fixed")
spec("taxi_premium",("d",0.3),"RT",-15000,-120000,["카카오T 블랙", "우버 블랙", "타다 플러스", "카카오T 벤티"],["택시","의전 차량","공항 이동"],"transfer",9,scale="fixed")
spec("shopping",("d",0.1),ALL,-15000,-350000,["(주)무신사", "쿠팡(주)", "네이버페이", "올리브영 강남타운점", "유니클로 명동중앙점", "자라 코엑스몰점", "다이소 신림점", "지그재그", "29CM"],["의류 구매","온라인 쇼핑","화장품","생활용품","신발","가방"],"transfer",21,scale=3)
spec("shopping_lux",("d",0.05),"RT",-1200000,-48000000,["샤넬 청담", "에르메스 메종 도산파크", "루이비통 메종 서울", "롤렉스 롯데 에비뉴엘", "까르띠에 신세계 강남", "디올 성수", "구찌 가옥"],["가방 구매","시계 구매","주얼리","의류","선물용 구매"],"transfer",15,scale="fixed")
spec("electronics",("d",0.012),ALL,-60000,-2800000,["Apple 가로수길", "삼성 디지털프라자 강남본점", "롯데하이마트 용산점", "쿠팡(주)", "컴퓨존"],["아이폰 구매","노트북","에어팟","TV 교체","세탁기 교체","냉장고","모니터"],"transfer",14,scale=2)
spec("repair_home",("d",0.01),"WMRT",-30000,-900000,["숨고", "귀뚜라미 보일러 서비스", "LG전자 서비스센터", "삼성전자 서비스센터", "서초설비"],["보일러 수리","세탁기 수리","수도 누수 수리","도어락 교체","에어컨 설치","전기 수리"],"transfer",15,scale=3)
spec("beauty",("d",0.05),ALL,-12000,-250000,["준오헤어 압구정점", "이철헤어커커 신사점", "네일리스트 서초점", "블루클럽 사당점", "라온 에스테틱"],["미용실","커트","펌","염색","네일","피부관리"],"transfer",14,scale=3)
spec("culture",("d",0.07),ALL,-9000,-180000,["CGV 용산아이파크몰", "메가박스 코엑스", "롯데시네마 월드타워", "(주)인터파크", "예스24", "교보문고 홍대입구점", "교보문고 광화문점", "멜론티켓"],["영화","콘서트 티켓","뮤지컬","도서 구매","전시회","공연"],"transfer",19,scale=1.5)
spec("hobby",("d",0.04),"WMRT",-20000,-600000,["골프존파크 강남점", "스크린골프 서초", "서초볼링센터", "더클라임 양재점", "피싱마트 낚시용품", "코베아 캠핑스토어", "양재 테니스장"],["취미 활동","레슨비","장비 구매","라운딩","레슨"],"transfer",18,scale=2)
spec("golf",("d",0.1),"RT",-250000,-1800000,["남서울CC","레이크사이드CC","해슬리나인브릿지","잭니클라우스GC"],["라운딩 그린피","캐디피","회원권 연회비 분납","골프 회식"],"transfer",13,scale="fixed")
spec("travel",("d",0.008),"WMRT",-350000,-4500000,["대한항공", "아시아나항공", "제주항공", "아고다", "야놀자", "호텔스컴바인", "하나투어"],["항공권","숙박","패키지 여행","호텔 예약","제주 여행","렌터카"],"transfer",10,scale=3)
spec("travel_lux",("d",0.01),"RT",-5000000,-85000000,["싱가포르항공 퍼스트","에미레이트 퍼스트","아만 리조트","포시즌스","리츠칼튼","프라이빗 제트"],["퍼스트클래스 항공권","리조트 숙박","전세기","럭셔리 크루즈","풀빌라"],"transfer",10,scale="fixed")
spec("kids",("d",0.05),"MR",-15000,-480000,["YBM 영어유치원", "챔피언 키즈카페", "구몬학습", "국기원 태권도장", "야마하 피아노학원", "CMS 수학학원", "청담어학원", "토이저러스 잠실점"],["학원비","교재비","학습지","아이 옷","장난감","체험학습비"],"transfer",16,scale=3)
spec("kids_private",("m",5),"RT",-1500000,-8000000,["국제학교","유학원","과외 선생님","승마장"],["국제학교 학비 분납","과외비","유학 송금","승마 레슨"],"transfer",10,scale="fixed")
spec("pet",("d",0.04),"MRT",-12000,-180000,["펫프렌즈", "어바웃펫", "멍멍살롱 애견미용", "강남 펫호텔"],["사료","간식","미용","호텔링","용품"],"transfer",15,scale=2)
spec("donation",("m",28),"MRT",-10000,-300000,["유니세프","굿네이버스","세이브더칠드런","월드비전","초록우산"],["정기 후원","후원금","기부"],"autoTransfer",10)
spec("religion",("d",0.14),"WMRT",-10000,-200000,["OO교회","OO성당","OO사"],["헌금","십일조","봉헌","시주"],"transfer",11,scale=5)
spec("edu_self",("d",0.01),"WMRT",-50000,-900000,["패스트캠퍼스","인프런","클래스101","YBM 어학원","해커스"],["온라인 강의","어학 수강","자격증 응시료","학원 등록"],"transfer",20,scale=2)
spec("laundry",("d",0.06),ALL,-3000,-45000,["크린토피아 서초점", "런드리고", "화이트세탁소"],["세탁","드라이클리닝","이불 세탁"],"transfer",17,scale="fixed")
spec("household_help",("m",4),"RT",-1800000,-6500000,["가사도우미","입주 도우미","운전기사 급여","정원 관리"],["가사도우미 월급","기사 급여","입주 도우미 급여","가정부 월급"],"transfer",10,scale="fixed")

# ---------- 돌발 ----------
spec("lost_phone",("d",0.0015),ALL,-400000,-1900000,["Apple 가로수길", "삼성전자 서비스센터", "KT 플라자"],["휴대폰 분실 재구매","액정 수리","침수 수리","분실 보험 자기부담금"],"transfer",16,scale="fixed")
spec("scam",("d",0.0008),"BW",-200000,-3000000,["보이스피싱 피해","스미싱 결제","중고거래 사기"],["사기 피해","피싱 피해금","사기 송금"],"transfer",15,scale="fixed")
spec("found_money",("d",0.002),ALL,10000,300000,["휴면계좌 찾기","미환급금","건강보험 환급","통신비 환급"],["휴면예금 찾기","미수령 환급금","건보료 환급","통신비 과오납 환급"],"transfer",11,scale="fixed")
spec("inherit",("d",0.0004),"MRT",5000000,300000000,["상속 재산","유산 분배","외할머니 유산"],["상속금 입금","유산 분배금","증여"],"transfer",14)
spec("legal",("d",0.001),"MRT",-500000,-12000000,["법무법인","김앤장","법률사무소"],["변호사 수임료","소송 비용","법률 자문료","공증 비용"],"transfer",11)
spec("moving",("d",0.0015),"WMRT",-300000,-2500000,["YES2404 이사", "통인익스프레스", "용달이사"],["이사 비용","포장이사","입주 청소"],"transfer",9,scale=3)
spec("deposit_return",("d",0.0012),"BWM",300000,20000000,["집주인","부동산"],["보증금 반환","월세 보증금 돌려받음"],"transfer",14,scale="fixed")
spec("deposit_pay",("d",0.0012),"BWM",-300000,-20000000,["집주인","부동산"],["보증금 납부","계약금","잔금"],"transfer",14,scale="fixed")
spec("art_buy",("d",0.004),"T",-30000000,-900000000,["국제갤러리","소더비","크리스티","갤러리현대","서울옥션"],["미술품 구매","경매 낙찰","조각 작품 구매"],"transfer",15,scale="fixed")
spec("yacht",("m",3),"T",-18000000,-60000000,["마리나 관리사","요트 관리 업체"],["요트 정박료","요트 관리비","선원 급여"],"autoTransfer",10,scale="fixed")
spec("realestate_buy",("d",0.0015),"T",-800000000,-12000000000,["부동산 매도인","법무사","신탁사"],["부동산 매입 잔금","건물 매입","토지 매입 계약금"],"transfer",11,scale="fixed")
spec("realestate_sell",("d",0.0012),"T",900000000,15000000000,["부동산 매수인","법무사"],["부동산 매각 대금","건물 매각 잔금","토지 매각"],"transfer",11,scale="fixed")
spec("foundation",("m",15),"T",-50000000,-300000000,["VX재단","사회복지공동모금회"],["재단 출연","기부금 출연","장학금 기탁"],"autoTransfer",14,scale="fixed")
spec("jet",("d",0.08),"T",-20000000,-70000000,["전용기 운영사","VistaJet","NetJets"],["전세기 운항","전용기 유지비","조종사 급여"],"transfer",12,scale="fixed")
spec("staff_bonus",("y",12,24),"RT",-3000000,-80000000,["직원 보너스","비서실","가사 직원"],["직원 연말 보너스","성과급 지급","명절 보너스 지급"],"transfer",10,scale="fixed")
spec("alimony",("m",1),"R",-2000000,-8000000,["전 배우자"],["양육비 송금","위자료 분납"],"autoTransfer",9,scale="fixed")

# ---------- 거지 특유 ----------
spec("broke_work",("d",0.3),"B",8000,95000,["일용직 인력사무소","쿠팡 물류센터","전단지 알바","새벽 인력시장","택배 상하차","대리운전"],["일당","일용직 급여","물류 일당","대리 수입","전단지 알바비"],"transfer",19,scale="fixed")
spec("broke_sell",("d",0.03),"B",5000,80000,["당근마켓","전당포","고물상","헌책방"],["물건 팔음","전당포 담보 대출","고물 판매","책 판매"],"transfer",16,scale="fixed")
spec("broke_fail",("d",0.08),"B",-1000,-5000,["은행","카드사","통신사"],["자동이체 실패 수수료","잔액부족 수수료","미납 안내 수수료"],"interest",6,scale="fixed")
spec("broke_cut",("d",0.01),"B",-20000,-90000,["한전 재연결","도시가스 재공급","통신사 재개통"],["단전 재연결비","가스 재공급비","휴대폰 정지 해제"],"transfer",10,scale="fixed")
spec("broke_seize",("d",0.01),"B",-50000,-400000,["법원 집행관","채권추심회사","압류"],["압류","급여 압류","통장 압류 집행","가압류"],"transfer",9,scale="fixed")
spec("broke_help",("d",0.02),"B",20000,300000,["동사무소","주민센터","긴급복지","복지재단","교회 구제비"],["긴급생계비","긴급복지 지원금","구제비","쌀 지원 대신 현금","기초생활 수급비"],"transfer",10,scale="fixed")
spec("broke_welfare",("m",20),"B",300000,650000,["국민기초생활보장","주민센터"],["기초생활 수급비","생계급여","주거급여"],"transfer",9,scale="fixed")
spec("broke_lend",("d",0.015),"B",-10000,-50000,["친구","전 직장 동료","동생"],["빌린 돈 일부 갚음","밥값 갚음"],"transfer",21,scale="fixed")
spec("broke_game",("d",0.04),"B",-5000,-60000,["모바일 게임","스포츠토토","경마","홀덤펍"],["현질","토토","경마","도박"],"transfer",23,scale="fixed")

# ---------- 서민 특유 ----------
spec("working_ot_meal",("d",0.2),"W",-7000,-14000,["회사 구내식당", "김가네 선릉점", "GS25 삼성역점"],["구내식당","야근 식대","점심 도시락"],"transfer",12,scale="fixed")
spec("working_hwe",("d",0.06),"W",-20000,-80000,["회식 N빵","동호회 회비","동창회"],["회식비","모임 회비","동창회비","N빵 송금"],"transfer",21,scale="fixed")
spec("working_saving",("m",26),"W",-100000,-300000,["{bank} 자유적금", "청년도약계좌", "주택청약"],["적금 자동이체","청약 납입","청년도약계좌"],"autoTransfer",9,scale="fixed")
spec("working_parent",("m",27),"W",-200000,-400000,["엄마","아빠","부모님"],["생활비 보태드림","부모님 용돈"],"transfer",10,scale="fixed")

import json, collections
counts = collections.Counter()
for s in S:
    for t in s["tiers"]: counts[t] += len(s["cps"])*len(s["memos"])
print("variations per tier:", dict(counts), "specs:", len(S))
assert all(counts[t] >= 1000 for t in ALL), "need >=1000 per tier"
tiermap={"B":".broke","W":".working","M":".middle","R":".rich","T":".tycoon"}
def sw(x): return json.dumps(x, ensure_ascii=False)
lines=["import Foundation","","/// Generated by scripts/gen_events.py — one-off and recurring life events. Do not edit by hand.",
"/// Amounts are middle-class KRW baselines; EventLibrary.tierScale adjusts them per tier.",
"struct LifeEventSpec {","    enum Kind { case daily(Double), monthly(Int), yearly(Int, Int) }",
"    let id: String","    let kind: Kind","    let tiers: [WealthTier]","    let baseMin: Double","    let baseMax: Double",
"    let category: TransactionCategory","    let hour: Int","    let counterparties: [String]","    let memos: [String]",
"    var scaling: Scaling = .full",
"    var variations: Int { counterparties.count * memos.count }","",
"    /// How the middle-class base amount follows the wealth tier.",
"    enum Scaling {",
"        case fixed            // same price for everyone (편의점, 택시, 약국) or base already at tier level",
"        case capped(Double)   // richer people pick pricier options, but only up to this multiple",
"        case full             // income, taxes, assets — scale all the way",
"        func factor(_ tierScale: Double) -> Double {",
"            switch self {",
"            case .fixed: return 1",
"            case .capped(let cap): return min(max(tierScale, 0.6), cap)",
"            case .full: return tierScale",
"            }",
"        }",
"    }",
"}","","enum EventLibrary {",
"    static let tierScale: [WealthTier: Double] = [.broke: 0.18, .working: 0.55, .middle: 1.0, .rich: 7.0, .tycoon: 90.0]","",
"    static func specs(for tier: WealthTier) -> [LifeEventSpec] { all.filter { $0.tiers.contains(tier) } }",
"    static func variationCount(for tier: WealthTier) -> Int { specs(for: tier).reduce(0) { $0 + $1.variations } }","",
"    static let all: [LifeEventSpec] = ["]
for s in S:
    k=s["kind"]
    kind = f".daily({k[1]})" if k[0]=="d" else (f".monthly({k[1]})" if k[0]=="m" else f".yearly({k[1]}, {k[2]})")
    tiers="["+", ".join(tiermap[t] for t in s["tiers"])+"]"
    lines.append(f'        LifeEventSpec(id: {sw(s["id"])}, kind: {kind}, tiers: {tiers}, baseMin: {s["lo"]}, baseMax: {s["hi"]}, category: .{s["cat"]}, hour: {s["hour"]},')
    lines.append(f'                      counterparties: {sw(s["cps"])},')
    sc=s["scale"]
    scaling = "" if sc=="full" else (", scaling: .fixed" if sc=="fixed" else f", scaling: .capped({sc})")
    lines.append(f'                      memos: {sw(s["memos"])}{scaling}),')
lines += ["    ]","}",""]
open("Vaultly/Models/EventLibrary.swift","w").write("\n".join(lines))
print("written")
