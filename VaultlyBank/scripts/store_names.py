"""App Store variant: real company / brand names → fictional ones (guideline 5.2.1).

One mapping, two outputs:
  1. Vaultly/App/StoreNames.swift — runtime filter used by Ledger.apply (every transaction),
     Brand (user-typed names), the transfer bank list and notifications. Identity in personal builds.
  2. QuestLibrary.swift — gets an `#if STORE` copy of the quest list with every string literal mapped.

Run from VaultlyBank/:  python3 scripts/store_names.py
gen_quests.py calls dualize_quests() after writing QuestLibrary.swift, so regenerating keeps the store copy.
Mode: 0 = plain substring, 1 = Latin token (no letter on either side), 2 = regular expression.
"""
import json, re, sys, os

MAP = [
    # ── conglomerates & manufacturers ──
    ("삼성바이오로직스", "성운바이오로직스", 0), ("삼성전자", "성운전자", 0), ("삼성생명", "성운생명", 0),
    ("삼성화재", "성운화재", 0), ("삼성서울병원", "성운서울병원", 0), ("강북삼성병원", "강북성운병원", 0),
    ("삼성 디지털프라자", "성운 디지털스토어", 0), (r"삼성(?!역|중앙역|동)", "성운", 2),
    ("현대자동차", "한결모터스", 0), ("현대차", "한결모터스", 0), ("현대해상", "한결해상", 0),
    ("현대카드", "한결카드", 0), ("현대캐피탈", "한결캐피탈", 0), ("현대오일뱅크", "한결오일", 0),
    ("현대 블루핸즈", "한결 정비센터", 0), ("갤러리현대", "갤러리한결", 0), (r"현대(?!인|적|사회|미술|식)", "한결", 2),
    ("기아 오토큐", "기린 정비센터", 0),
    ("LG전자", "엘브릿지전자", 0), ("LG유플러스", "유플넷", 0), ("LG", "LN", 1),
    ("SK하이닉스", "SN반도체", 0), ("하이닉스", "SN반도체", 0), ("SK텔레콤", "SN텔레콤", 0),
    ("SK브로드밴드", "SN브로드밴드", 0), ("SK에너지", "SN에너지", 0), ("SK", "SN", 1),
    ("한화에어로스페이스", "한빛에어로", 0), ("한화생명", "한빛생명", 0), ("한화손해보험", "한빛손해보험", 0), ("한화", "한빛", 0),
    ("CJ제일제당", "CN푸드", 0), ("CJ대한통운", "한빛통운", 0), ("대한통운", "한빛통운", 0), ("CJ", "CN", 1),
    ("신세계", "신하늘", 0), ("코오롱모터스", "코모모터스", 0), ("한성자동차", "한별자동차", 0),
    ("코카콜라", "코크앤코", 0), ("테슬라", "테슬린", 0),
    # ── retail, convenience, food ──
    ("GS25", "가온25", 0), ("GS더프레시", "가온프레시", 0), ("GS칼텍스", "가온오일", 0), ("GS", "가온", 1),
    ("CU", "큐마트", 1), ("세븐일레븐", "일레븐마트", 0), ("이마트24", "데이마트24", 0), ("이마트 에브리데이", "데이마트 에브리데이", 0),
    ("이마트", "데이마트", 0), ("롯데리아", "로하버거", 0), ("롯데마트", "로하마트", 0), ("롯데카드", "로하카드", 0),
    ("롯데손해보험", "로하손해보험", 0), ("롯데하이마트", "로하전자", 0), ("롯데시네마", "로하시네마", 0), ("롯데", "로하", 0),
    ("홈플러스", "홈마켓", 0), ("코스트코", "메가창고", 0), ("노브랜드", "심플마켓", 0), ("다이소", "천원샵", 0),
    ("스타벅스", "별다방", 0), ("이디야커피", "이데아커피", 0), ("이디야", "이데아", 0), ("투썸플레이스", "투뷰플레이스", 0), ("투썸", "투뷰", 0),
    ("맥도날드", "맥버거", 0), ("버거킹", "버거왕", 0), ("맘스터치", "맘스버거", 0), ("BBQ", "BBX", 1),
    ("교촌치킨", "교동치킨", 0), ("서브웨이", "서브샌드", 0), ("본죽", "본가죽", 0), ("김밥천국", "김밥나라", 0),
    ("한솥도시락", "한끼도시락", 0), ("김가네", "김씨네", 0), ("메가커피", "기가커피", 0), ("빽다방", "빡다방", 0),
    ("올리브영", "그린영", 0), ("무신사", "무드샵", 0), ("29CM", "30CM", 0), ("지그재그", "지그잭", 0),
    ("11번가", "12번가", 0), ("G마켓", "H마켓", 0), ("인터파크", "인터랜드", 0), ("예스24", "북스24", 0),
    ("교보문고", "한글문고", 0), ("교보생명", "한글생명", 0), ("교보", "한글", 0),
    ("유니클로", "유니폼샵", 0), ("자라 코엑스몰점", "자르 코엑스몰점", 0), ("토이저러스", "토이랜드", 0),
    ("컴퓨존", "컴퓨랜드", 0), ("코베아", "코브라캠프", 0), ("타이어뱅크", "타이어마트", 0), ("스피드메이트", "스피드케어", 0),
    # ── platforms ──
    ("쿠팡이츠", "로켓몰이츠", 0), ("쿠팡플렉스", "로켓몰플렉스", 0), ("쿠팡", "로켓몰", 0),
    ("우아한형제들", "배달형제들", 0), ("배민커넥트", "배달형제 커넥트", 0), ("배민", "배달형제", 0),
    ("요기요", "여기요", 0), ("당근마켓", "오이마켓", 0), (r"당근(?!케이크|주스)", "오이마켓", 2), ("번개장터", "번쩍장터", 0),
    ("야놀자", "놀자고", 0), ("아고다", "아고라트립", 0), ("호텔스컴바인", "호텔모아", 0), ("하나투어", "두나투어", 0),
    ("에어비앤비", "에어스테이", 0), ("숨고", "고수찾기", 0), ("크몽", "크림몽", 0), ("클래스101", "클래스102", 0),
    ("패스트캠퍼스", "패스트스쿨", 0), ("인프런", "인런", 0), ("알바천국", "알바나라", 0), ("런드리고", "런드리온", 0),
    ("크린토피아", "크린월드", 0), ("네이버웹툰", "누리웹툰", 0), ("네이버페이", "누리페이", 0), ("네이버", "누리넷", 0),
    ("카카오뱅크", "코코뱅크", 0), ("카카오페이", "코코페이", 0), ("카카오T", "코코T", 0), ("카카오", "코코", 0),
    ("토스뱅크", "미니멀뱅크", 0), ("토스증권", "미니멀증권", 0), ("토스모바일", "미니멀모바일", 0), (r"토스(?!트)", "미니멀페이", 2),
    ("페이코", "페이온", 0), ("티머니", "티캐시", 0), ("업비트", "업코인", 0), ("빗썸", "코인마루", 0),
    ("유튜브", "뷰튜브", 0), ("넷플릭스", "넷무비", 0), ("디즈니+", "판타지+", 0), ("티빙", "티비온", 0), ("왓챠", "왓치온", 0),
    ("웨이브", "웨이브온", 0), ("멜론", "멜로디", 0), ("지니뮤직", "지니어스뮤직", 0), ("스포티파이", "스포티온", 0),
    ("애플뮤직", "페어뮤직", 0), ("애플", "페어", 0), ("Apple", "Pear", 1), ("App Store", "App Market", 0), ("iCloud", "myCloud", 0),
    ("Google", "Gooble", 1), ("Dropbox", "Dropkit", 1), ("Steam", "Vapor", 1), ("닌텐도", "게임닷", 0),
    ("아이폰", "페어폰", 0), ("맥북", "페어북", 0), ("CGV", "CGX", 1), ("메가박스", "메가씨네", 0),
    ("골프존", "골프랜드", 0), ("YBM", "YBN", 1), ("해커스", "해커즈", 0), ("구몬학습", "구문학습", 0), ("야마하", "야마노", 0),
    ("YES2404", "YES2000", 0), ("통인익스프레스", "통일익스프레스", 0),
    # ── banks, cards, insurers, lenders ──
    ("KB국민은행", "온국민은행", 0), ("KB손해보험", "KN손해보험", 0), ("KB금융", "KN금융", 0), ("KB", "KN", 1),
    ("신한은행", "신화은행", 0), ("신한라이프", "신화라이프", 0), ("신한카드", "신화카드", 0), ("신한", "신화", 0),
    ("우리은행", "마루은행", 0), ("우리캐피탈", "마루캐피탈", 0), ("우리카드", "마루카드", 0), ("JB", "JN", 1),
    ("하나은행", "두나은행", 0), ("하나캐피탈", "두나캐피탈", 0), ("하나카드", "두나카드", 0),
    ("케이뱅크", "제이뱅크", 0), ("IBK", "IBX", 1), ("NH농협", "NM농우", 0), ("농협", "농우", 0), ("NH", "NM", 1),
    ("미래에셋", "미래자산", 0), ("키움", "키운", 0), ("메리츠", "메리트", 0), ("흥국생명", "흥진생명", 0),
    ("라이나생명", "라임생명", 0), ("라이나", "라임", 0), ("AIA", "AIX", 1), ("AIG", "AGX", 1), ("AXA", "AZA", 1), ("ABL", "ABX", 1),
    ("DB손해보험", "DN손해보험", 0), ("DB손보", "DN손보", 0), ("DB 펫블리", "DN 펫블리", 0),
    ("동양생명", "동해생명", 0), ("처브라이프", "체브라이프", 0), ("러시앤캐시", "러닝앤캐시", 0), ("리드코프", "리드론", 0),
    ("OK캐피탈", "OX캐피탈", 0), ("BMW파이낸셜", "BMX파이낸셜", 0),
    ("KT M모바일", "QT M모바일", 0), ("KT", "QT", 1), ("딜라이브", "딜라인", 0),
    # ── cars, luxury, travel, property ──
    ("BMW", "BMX", 1), ("벤츠", "멘츠", 0), ("포르쉐", "포르젠", 0), ("람보르기니", "람보르노", 0), ("아반떼", "아르떼", 0),
    ("롤렉스", "로렉시아", 0), ("에르메스", "에르미스", 0), ("샤넬", "샤르넬", 0), ("루이비통", "루이벨롱", 0),
    ("구찌", "구띠", 0), ("디올", "디오렌", 0), ("까르띠에", "까르티아", 0),
    ("대한항공", "한국에어", 0), ("아시아나항공", "아시아에어", 0), ("아시아나", "아시아에어", 0), ("제주항공", "탐라항공", 0),
    ("싱가포르항공", "라이온항공", 0), ("에미레이트", "데저트항공", 0), ("NetJets", "SkyShare Jets", 1), ("VistaJet", "AzureJet", 1),
    ("시그니엘", "시그나엘", 0), ("리츠칼튼", "리츠칼라", 0), ("포시즌스", "식스시즌스", 0), ("아만 리조트", "아마노 리조트", 0),
    ("래미안퍼스티지", "래미온퍼스트", 0), ("래미안", "래미온", 0), ("아크로리버파크", "아크라리버파크", 0),
    ("자이 관리", "자인 관리", 0), ("힐스테이트", "힐스타운", 0), ("LH", "KH", 1),
    # ── health, law, auctions, golf, charity, lottery ──
    ("서울아산병원", "서울한마음병원", 0), ("아산병원", "한마음병원", 0), ("서울대병원", "서울국립병원", 0),
    ("세브란스", "세브린", 0), ("차움", "차온", 0), ("하나로의료재단", "하나로검진재단", 0), ("KMI한국의학연구소", "KMX건강검진센터", 0),
    ("김앤장", "김앤리", 0), ("서울옥션", "서울경매", 0), ("소더비", "소더스", 0), ("크리스티", "크리스톤", 0),
    ("남서울CC", "남한울CC", 0), ("레이크사이드CC", "레이크힐CC", 0), ("잭니클라우스", "잭니콜스", 0), ("해슬리나인브릿지", "나인힐스", 0),
    ("굿네이버스", "굿이웃", 0), ("월드비전", "월드호프", 0), ("유니세프", "유니케어", 0), ("세이브더칠드런", "세이브키즈", 0),
    ("초록우산", "초록나무", 0), ("동행복권", "행운복권", 0), ("스포츠토토", "스포츠복권", 0), ("스피또", "스피드복권", 0),
    ("하버드", "하버포드", 0), ("준오헤어", "준호헤어", 0), ("이철헤어커커", "이수헤어", 0),
    ("라연", "라온", 0), ("밍글스", "밍글", 0), ("정식당", "정갈당", 0), ("스시조", "스시도", 0), ("스시 효", "스시 하루", 0),
]

def _ordered():
    # longest literal name first so 삼성전자 wins over 삼성 and GS25 over GS; the generic regex rules
    # (bare 삼성 / 현대 / 토스 / 당근) always run last
    return sorted(MAP, key=lambda t: (t[2] == 2, -len(t[0])))

def _compiled():
    out = []
    for real, fake, mode in _ordered():
        if mode == 0:
            out.append((re.compile(re.escape(real)), fake))
        elif mode == 1:
            out.append((re.compile(r"(?<![A-Za-z])" + re.escape(real) + r"(?![A-Za-z])"), fake))
        else:
            out.append((re.compile(real), fake))
    return out

def fix(s, _rx=[]):
    if not _rx:
        _rx.extend(_compiled())
    # one pass per rule, but never re-touch text an earlier rule produced: mark replacements with \x00
    for rx, fake in _rx:
        s = rx.sub(lambda m: "\x00" + fake + "\x01", s)
    return re.sub(r"[\x00\x01]", "", s)

def swift_table():
    def lit(x): return json.dumps(x, ensure_ascii=False)
    rows = [f"        ({lit(r)}, {lit(f)}, {m})," for r, f, m in _ordered()]
    return "\n".join([
        "import Foundation", "",
        "/// Generated by scripts/store_names.py — do not edit by hand.",
        "/// App Store build only: real company and brand names become fictional ones (guideline 5.2.1),",
        "/// so the store variant never shows a real bank, merchant, insurer or employer. Identity in personal builds.",
        "enum StoreNames {",
        "    /// (real, fictional, mode) — mode 0 plain, 1 Latin token, 2 regular expression. Longest first.",
        "    static let pairs: [(String, String, Int)] = [",
        *rows,
        "    ]", "",
        "    private static let lock = NSLock()",
        "    nonisolated(unsafe) private static var cache: [String: String] = [:]", "",
        "    /// Fictional names for the transfer screen's bank picker.",
        "    static let banks = [\"코코뱅크\", \"미니멀뱅크\", \"제이뱅크\", \"온국민은행\", \"신화은행\", \"마루은행\", \"두나은행\", \"NM농우은행\", \"IBX기업은행\"]", "",
        "    static func fix(_ s: String) -> String {",
        "        guard AppVariant.isStore, !s.isEmpty else { return s }",
        "        lock.lock(); defer { lock.unlock() }",
        "        if let hit = cache[s] { return hit }",
        "        var out = s",
        "        for (real, fake, mode) in pairs {",
        "            let pattern: String",
        "            switch mode {",
        "            case 0:",
        "                guard out.contains(real) else { continue }",
        "                pattern = NSRegularExpression.escapedPattern(for: real)",
        "            case 1:",
        "                guard out.contains(real) else { continue }",
        "                pattern = \"(?<![A-Za-z])\" + NSRegularExpression.escapedPattern(for: real) + \"(?![A-Za-z])\"",
        "            default:",
        "                pattern = real",
        "            }",
        "            // \\u{0} … \\u{1} fence each replacement so a later, shorter rule can't rewrite it again",
        "            out = out.replacingOccurrences(of: pattern, with: \"\\u{0}\" + NSRegularExpression.escapedTemplate(for: fake) + \"\\u{1}\",",
        "                                           options: .regularExpression)",
        "        }",
        "        out = out.replacingOccurrences(of: \"\\u{0}\", with: \"\").replacingOccurrences(of: \"\\u{1}\", with: \"\")",
        "        if cache.count > 4000 { cache.removeAll() }",
        "        cache[s] = out",
        "        return out",
        "    }",
        "}", ""])

STR = re.compile(r'"(?:[^"\\]|\\.)*"')

def dualize_quests(path="Vaultly/Models/QuestLibrary.swift"):
    src = open(path, encoding="utf-8").read()
    if "#if STORE" in src:
        # rebuild from the personal copy
        head, rest = src.split("#if STORE\n", 1)
        personal = rest.split("#else\n", 1)[1].split("#endif\n", 1)[0]
        tail = rest.split("#endif\n", 1)[1]
        src = head + personal + tail
    start = src.index("    // MARK: quests\n")
    end = src.index("    // MARK: end quests\n") + len("    // MARK: end quests\n")
    block = src[start:end]
    store_block = STR.sub(lambda m: json.dumps(fix(json.loads(m.group(0))), ensure_ascii=False), block)
    out = src[:start] + "#if STORE\n" + store_block + "#else\n" + block + "#endif\n" + src[end:]
    open(path, "w", encoding="utf-8").write(out)

if __name__ == "__main__":
    os.chdir(os.path.join(os.path.dirname(__file__), ".."))
    open("Vaultly/App/StoreNames.swift", "w", encoding="utf-8").write(swift_table())
    dualize_quests()
    print("StoreNames.swift written, QuestLibrary.swift dualized")
