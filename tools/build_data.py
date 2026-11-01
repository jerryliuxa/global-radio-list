# -*- coding: utf-8 -*-
"""
build_data.py — 把 xta26.xlsx（A26 短波频率表）解析为前端可用的 data.js

用法: python tools/build_data.py [xlsx路径] [输出目录]
默认: python tools/build_data.py "C:/Users/Administrator/WorkBuddy/短波频率表/xta26.xlsx" ../data
"""
import sys, os, re, json, collections, datetime

import openpyxl

# ---------------- 国家/地区代码映射 (代码 -> (中文名, 英文名, ISO2 用于旗帜)) ----------------
COUNTRIES = {
    'ALG': ('阿尔及利亚', 'Algeria', 'DZ'), 'ALS': ('阿拉斯加(美)', 'Alaska (US)', 'US'),
    'ARM': ('亚美尼亚', 'Armenia', 'AM'), 'ARS': ('沙特阿拉伯', 'Saudi Arabia', 'SA'),
    'ASC': ('阿森松岛', 'Ascension', 'SH'), 'ATA': ('阿根廷', 'Argentina', 'AR'),
    'AUS': ('澳大利亚', 'Australia', 'AU'), 'B': ('巴西', 'Brazil', 'BR'),
    'BEL': ('比利时', 'Belgium', 'BE'), 'BEN': ('贝宁', 'Benin', 'BJ'),
    'BGD': ('孟加拉国', 'Bangladesh', 'BD'), 'BHR': ('巴林', 'Bahrain', 'BH'),
    'BOL': ('玻利维亚', 'Bolivia', 'BO'),
    'BOT': ('博茨瓦纳', 'Botswana', 'BW'), 'BTN': ('不丹', 'Bhutan', 'BT'),
    'BUL': ('保加利亚', 'Bulgaria', 'BG'), 'CAN': ('加拿大', 'Canada', 'CA'),
    'CHN': ('中国', 'China', 'CN'), 'CLM': ('哥伦比亚', 'Colombia', 'CO'),
    'CLN': ('斯里兰卡', 'Sri Lanka', 'LK'), 'CUB': ('古巴', 'Cuba', 'CU'),
    'CVA': ('梵蒂冈', 'Vatican', 'VA'), 'CYP': ('塞浦路斯', 'Cyprus', 'CY'),
    'CZE': ('捷克', 'Czechia', 'CZ'), 'D': ('德国', 'Germany', 'DE'),
    'DNK': ('丹麦', 'Denmark', 'DK'), 'E': ('西班牙', 'Spain', 'ES'),
    'EGY': ('埃及', 'Egypt', 'EG'), 'EQA': ('厄瓜多尔', 'Ecuador', 'EC'),
    'EST': ('爱沙尼亚', 'Estonia', 'EE'), 'ETH': ('埃塞俄比亚', 'Ethiopia', 'ET'),
    'F': ('法国', 'France', 'FR'), 'FIN': ('芬兰', 'Finland', 'FI'),
    'FRO': ('法罗群岛', 'Faroe Is.', 'FO'), 'G': ('英国', 'United Kingdom', 'GB'),
    'GRC': ('希腊', 'Greece', 'GR'), 'GUI': ('几内亚', 'Guinea', 'GN'),
    'GUM': ('关岛', 'Guam', 'GU'), 'HND': ('洪都拉斯', 'Honduras', 'HN'),
    'HNG': ('匈牙利', 'Hungary', 'HU'), 'HOL': ('荷兰', 'Netherlands', 'NL'),
    'I': ('意大利', 'Italy', 'IT'), 'IND': ('印度', 'India', 'IN'),
    'INS': ('印度尼西亚', 'Indonesia', 'ID'), 'IRL': ('爱尔兰', 'Ireland', 'IE'),
    'IRN': ('伊朗', 'Iran', 'IR'), 'ISL': ('冰岛', 'Iceland', 'IS'),
    'J': ('日本', 'Japan', 'JP'), 'KGZ': ('吉尔吉斯斯坦', 'Kyrgyzstan', 'KG'),
    'KOR': ('韩国', 'South Korea', 'KR'), 'KRE': ('朝鲜', 'North Korea', 'KP'),
    'KWT': ('科威特', 'Kuwait', 'KW'), 'LAO': ('老挝', 'Laos', 'LA'),
    'LBR': ('利比里亚', 'Liberia', 'LR'), 'LTU': ('立陶宛', 'Lithuania', 'LT'),
    'LUX': ('卢森堡', 'Luxembourg', 'LU'), 'MCO': ('摩纳哥', 'Monaco', 'MC'),
    'MDA': ('摩尔多瓦', 'Moldova', 'MD'), 'MDG': ('马达加斯加', 'Madagascar', 'MG'),
    'MEX': ('墨西哥', 'Mexico', 'MX'), 'MLA': ('马来西亚', 'Malaysia', 'MY'),
    'MLD': ('马尔代夫', 'Maldives', 'MV'), 'MLI': ('马里', 'Mali', 'ML'),
    'MMR': ('缅甸', 'Myanmar', 'MM'), 'MNG': ('蒙古', 'Mongolia', 'MN'),
    'MRC': ('摩洛哥', 'Morocco', 'MA'), 'NIG': ('尼日利亚', 'Nigeria', 'NG'),
    'NOR': ('挪威', 'Norway', 'NO'), 'NZL': ('新西兰', 'New Zealand', 'NZ'),
    'OMA': ('阿曼', 'Oman', 'OM'), 'PHL': ('菲律宾', 'Philippines', 'PH'),
    'PLW': ('帕劳', 'Palau', 'PW'), 'POL': ('波兰', 'Poland', 'PL'),
    'PRU': ('秘鲁', 'Peru', 'PE'), 'QAT': ('卡塔尔', 'Qatar', 'QA'),
    'ROU': ('罗马尼亚', 'Romania', 'RO'), 'RUS': ('俄罗斯', 'Russia', 'RU'),
    'SLM': ('所罗门群岛', 'Solomon Is.', 'SB'), 'SVK': ('斯洛伐克', 'Slovakia', 'SK'),
    'SVN': ('斯洛文尼亚', 'Slovenia', 'SI'), 'SWZ': ('斯威士兰', 'Eswatini', 'SZ'),
    'SYR': ('叙利亚', 'Syria', 'SY'), 'THA': ('泰国', 'Thailand', 'TH'),
    'TIB': ('中国西藏', 'Tibet (CN)', 'CN'), 'TJK': ('塔吉克斯坦', 'Tajikistan', 'TJ'),
    'TKS': ('中国新疆', 'Xinjiang (CN)', 'CN'), 'TUN': ('突尼斯', 'Tunisia', 'TN'),
    'TUR': ('土耳其', 'Turkey', 'TR'), 'TWN': ('中国台湾', 'Taiwan', 'TW'),
    'UAE': ('阿联酋', 'UAE', 'AE'), 'UGA': ('乌干达', 'Uganda', 'UG'),
    'UKR': ('乌克兰', 'Ukraine', 'UA'), 'URU': ('乌拉圭', 'Uruguay', 'UY'),
    'USA': ('美国', 'USA', 'US'), 'UZB': ('乌兹别克斯坦', 'Uzbekistan', 'UZ'),
    'VTN': ('越南', 'Vietnam', 'VN'), 'VUT': ('瓦努阿图', 'Vanuatu', 'VU'),
    'ZMB': ('赞比亚', 'Zambia', 'ZM'),
}

# ---------------- 语言归一化（原始写法/ISO 639 -> 统一分组键） ----------------
LANG_MAP = {
    'english': '英语', 'eng': '英语', 'engl': '英语', 'en': '英语',
    'chinese': '中文', 'chn': '中文', 'zho': '中文', 'cmn': '中文', 'zh': '中文', 'mandarin': '中文',
    'cantonese': '粤语', 'yue': '粤语',
    'japanese': '日语', 'jpn': '日语', 'ja': '日语',
    'korean': '韩语', 'kor': '韩语', 'ko': '韩语',
    'russian': '俄语', 'rus': '俄语', 'ru': '俄语',
    'arabic': '阿拉伯语', 'ara': '阿拉伯语', 'arb': '阿拉伯语', 'ar': '阿拉伯语',
    'spanish': '西班牙语', 'spa': '西班牙语', 'es': '西班牙语', 'castillian': '西班牙语',
    'french': '法语', 'fra': '法语', 'fre': '法语', 'fr': '法语',
    'german': '德语', 'deu': '德语', 'ger': '德语', 'de': '德语',
    'portuguese': '葡语', 'por': '葡语', 'pt': '葡语', 'brazilian portuguese': '葡语',
    'italian': '意语', 'ita': '意语', 'it': '意语',
    'hindi': '印地语', 'hin': '印地语', 'hi': '印地语',
    'urdu': '乌尔都语', 'urd': '乌尔都语',
    'bengali': '孟加拉语', 'ben': '孟加拉语', 'bangla': '孟加拉语',
    'tamil': '泰米尔语', 'tam': '泰米尔语',
    'telugu': '泰卢固语', 'tel': '泰卢固语',
    'malayalam': '马拉雅拉姆语', 'mal': '马拉雅拉姆语',
    'uyghur': '维吾尔语', 'uig': '维吾尔语', 'uighur': '维吾尔语', 'uyhur': '维吾尔语',
    'tibetan': '藏语', 'bod': '藏语', 'tib': '藏语',
    'mongolian': '蒙语', 'mon': '蒙语', 'mng': '蒙语', 'halh': '蒙语',
    'vietnamese': '越南语', 'vie': '越南语', 'vtn': '越南语', 'vi': '越南语',
    'thai': '泰语', 'tha': '泰语',
    'burmese': '缅甸语', 'mya': '缅甸语', 'bur': '缅甸语',
    'indonesian': '印尼语', 'ind': '印尼语', 'id': '印尼语',
    'malay': '马来语', 'msa': '马来语',
    'persian': '波斯语', 'prs': '波斯语', 'fas': '波斯语', 'farsi': '波斯语', 'per': '波斯语',
    'dari': '达里语', 'prs dari': '达里语',
    'pashto': '普什图语', 'pus': '普什图语', 'pst': '普什图语', 'pushto': '普什图语',
    'turkish': '土耳其语', 'tur': '土耳其语',
    'polish': '波兰语', 'pol': '波兰语',
    'romanian': '罗马尼亚语', 'ron': '罗马尼亚语', 'rum': '罗马尼亚语',
    'bulgarian': '保加利亚语', 'bul': '保加利亚语',
    'ukrainian': '乌克兰语', 'ukr': '乌克兰语',
    'belarusian': '白俄罗斯语', 'bel': '白俄罗斯语', 'byelorussian': '白俄罗斯语',
    'swahili': '斯瓦希里语', 'swa': '斯瓦希里语', 'swah': '斯瓦希里语',
    'hausa': '豪萨语', 'hau': '豪萨语',
    'somali': '索马里语', 'som': '索马里语',
    'amharic': '阿姆哈拉语', 'amh': '阿姆哈拉语', 'amharic)': '阿姆哈拉语',
    'oromo': '奥罗莫语', 'orm': '奥罗莫语', 'oromic': '奥罗莫语',
    'afar': '阿法尔语', 'aar': '阿法尔语', 'afar)': '阿法尔语',
    'tigrinya': '提格雷尼亚语', 'tir': '提格雷尼亚语',
    'ndebele': '恩德贝勒语', 'nde': '恩德贝勒语', 'nbl': '恩德贝勒语',
    'shona': '绍纳语', 'sna': '绍纳语',
    'zulu': '祖鲁语', 'zul': '祖鲁语',
    'sesotho': '塞索托语', 'sot': '塞索托语', 'sotho': '塞索托语',
    'quechua': '克丘亚语', 'que': '克丘亚语', 'quichua': '克丘亚语',
    'aymara': '艾马拉语', 'aym': '艾马拉语',
    'guarani': '瓜拉尼语', 'grn': '瓜拉尼语',
    'creole': '克里奥尔语', 'cri': '克里奥尔语', 'kriol': '克里奥尔语', 'patois': '克里奥尔语',
    'sinhala': '僧伽罗语', 'sin': '僧伽罗语', 'singhala': '僧伽罗语',
    'nepali': '尼泊尔语', 'nep': '尼泊尔语',
    'khmer': '高棉语', 'khm': '高棉语', 'cambodian': '高棉语',
    'lao': '老挝语', 'lao)': '老挝语',
    'cebuano': '宿务语', 'ceb': '宿务语',
    'tagalog': '他加禄语', 'tgl': '他加禄语', 'filipino': '菲律宾语',
    'uzbek': '乌兹别克语', 'uzb': '乌兹别克语',
    'kazakh': '哈萨克语', 'kaz': '哈萨克语', 'kazakh)': '哈萨克语',
    'kyrgyz': '吉尔吉斯语', 'kir': '吉尔吉斯语',
    'turkmen': '土库曼语', 'tuk': '土库曼语',
    'tajik': '塔吉克语', 'tgk': '塔吉克语',
    'azeri': '阿塞拜疆语', 'aze': '阿塞拜疆语',
    'georgian': '格鲁吉亚语', 'kat': '格鲁吉亚语', 'geo': '格鲁吉亚语',
    'armenian': '亚美尼亚语', 'hye': '亚美尼亚语', 'arm': '亚美尼亚语',
    'hebrew': '希伯来语', 'heb': '希伯来语',
    'greek': '希腊语', 'ell': '希腊语', 'gre': '希腊语',
    'dutch': '荷兰语', 'nld': '荷兰语', 'dut': '荷兰语', 'flemish': '荷兰语',
    'hungarian': '匈牙利语', 'hun': '匈牙利语',
    'czech': '捷克语', 'cze': '捷克语', 'ces': '捷克语',
    'slovak': '斯洛伐克语', 'slk': '斯洛伐克语', 'slo': '斯洛伐克语',
    'slovenian': '斯洛文尼亚语', 'slv': '斯洛文尼亚语',
    'croatian': '克罗地亚语', 'hrv': '克罗地亚语', 'cro': '克罗地亚语',
    'serbian': '塞尔维亚语', 'srp': '塞尔维亚语', 'ser': '塞尔维亚语',
    'bosnian': '波斯尼亚语', 'bos': '波斯尼亚语',
    'macedonian': '马其顿语', 'mkd': '马其顿语', 'mac': '马其顿语',
    'albanian': '阿尔巴尼亚语', 'sqi': '阿尔巴尼亚语', 'alb': '阿尔巴尼亚语',
    'lithuanian': '立陶宛语', 'lit': '立陶宛语',
    'latvian': '拉脱维亚语', 'lav': '拉脱维亚语', 'lavisian': '拉脱维亚语',
    'estonian': '爱沙尼亚语', 'est)': '爱沙尼亚语',
    'finnish': '芬兰语', 'fin)': '芬兰语',
    'swedish': '瑞典语', 'swe': '瑞典语',
    'norwegian': '挪威语', 'nor': '挪威语',
    'danish': '丹麦语', 'dan': '丹麦语',
    'icelandic': '冰岛语', 'isl': '冰岛语', 'ice': '冰岛语',
    'bislama': '比斯拉马语', 'bis': '比斯拉马语',
    'fijian': '斐济语', 'fij': '斐济语',
    'samoan': '萨摩亚语', 'smo': '萨摩亚语',
    'tok pisin': '巴布亚皮钦语', 'tpi': '巴布亚皮钦语', 'pidgin': '巴布亚皮钦语',
    'hiri motu': '希里莫图语', 'hmo': '希里莫图语',
    'kurdish': '库尔德语', 'kur': '库尔德语', 'kurdi': '库尔德语',
    'kashmiri': '克什米尔语', 'kas': '克什米尔语',
    'sindhi': '信德语', 'snd': '信德语',
    'gujarati': '古吉拉特语', 'guj': '古吉拉特语',
    'marathi': '马拉地语', 'mar': '马拉地语',
    'punjabi': '旁遮普语', 'pan': '旁遮普语',
    'sinhalese': '僧伽罗语', 'divehi': '迪维希语', 'dv': '迪维希语',
    'assamese': '阿萨姆语', 'asm': '阿萨姆语',
    'balochi': '俾路支语', 'bal': '俾路支语',
    'brahui': '布拉灰语', 'brh': '布拉灰语',
    'aromanian': '阿罗马尼亚语', 'rup': '阿罗马尼亚语',
    'mul': '多语种', 'multiple': '多语种', 'various': '多语种',
    'und': '未指定', 'unidentified': '未指定', 'unknown': '未指定',
    'a1b': '授时信号', 'pwm/pm': '授时信号', 'pwm': '授时信号', 'pm': '授时信号',
    'standard time': '授时信号', 'time signal': '授时信号', 'time signals': '授时信号', 'f3c': '授时信号',
    # ---- 二轮补充（来自实际数据中未映射的语言） ----
    'hakka': '客家话', 'hak': '客家话',
    'amoy': '闽南语', 'nan': '闽南语', 'min nan': '闽南语', 'hokkien': '闽南语',
    'chaozhou': '潮州话', 'swatow': '潮州话',
    'fil': '菲律宾语', 'filipino': '菲律宾语', 'tagalog': '菲律宾语',
    'esperanto': '世界语', 'epo': '世界语',
    'belarussian': '白俄罗斯语', 'laotian': '老挝语',
    'kin': '卢旺达语', 'kinyarwanda': '卢旺达语', 'kirundi': '基隆迪语',
    'malaysian': '马来语', 'mly': '马来语',
    'javanese': '爪哇语', 'jv': '爪哇语',
    'kanuri': '卡努里语', 'malagasy': '马尔加什语', 'plt': '马尔加什语',
    'music': '音乐', 'jamming': '干扰信号', 'qpsk': '数字数据', 'nldeng': '荷兰语',
    'engspa': '英语', 'vernacular': '当地语言',
    'iu mien': '勉语（瑶族）', 'mien': '勉语（瑶族）',
    'darfur arabic': '阿拉伯语', 'juba arabic': '阿拉伯语', 'arabic.': '阿拉伯语',
    'sephardic': '拉迪诺语', 'ladino': '拉迪诺语',
    'mandinka': '曼丁哥语', 'mandingo': '曼丁哥语',
    'ful': '富拉语', 'fulfulde': '富拉语', 'fula': '富拉语', 'peul': '富拉语',
    'man': '满语', 'manchu': '满语',
    'dzo': '宗卡语', 'dzongkha': '宗卡语',
    'bengal': '孟加拉语', 'afan oromo': '奥罗莫语', 'oriya': '奥里亚语',
    'chattisgarhi': '恰蒂斯加尔语', 'chhattisgarhi': '恰蒂斯加尔语',
    'baluchi': '俾路支语', 'hmong': '苗语', 'miao': '苗语',
    'lisu': '傈僳语', 'lahu': '拉祜语',
    'shan': '掸语', 'shn': '掸语', 'rki': '若开语', 'rakhine': '若开语',
    'ksw': '克伦语', 'karen': '克伦语', 'rohingya': '罗兴亚语',
    'kannada': '卡纳达语', 'kan': '卡纳达语',
    'lat': '拉脱维亚语', 'bam': '班巴拉语', 'bambara': '班巴拉语',
    'latin': '拉丁语', 'wolof': '沃洛夫语', 'wol': '沃洛夫语',
    'twi': '契维语', 'ewe': '埃维语', 'susu': '苏苏语', 'krio': '克里奥尔语',
    'yoruba': '约鲁巴语', 'yor': '约鲁巴语', 'ibo': '伊博语', 'igbo': '伊博语',
    'efik': '埃菲克语', 'mossi': '莫西语', 'more': '莫西语',
    'lingala': '林加拉语', 'lin': '林加拉语', 'sango': '桑戈语', 'sgn': '桑戈语',
    'kikongo': '刚果语', 'kon': '刚果语', 'chokwe': '乔克韦语',
    'kimbundu': '金本杜语', 'umbundu': '姆本杜语',
    'lozi': '洛齐语', 'tumbuka': '图姆布卡语', 'chewa': '切瓦语', 'chichewa': '切瓦语',
    'nya': '切瓦语', 'tsonga': '聪加语', 'ts': '聪加语', 'venda': '文达语',
    'tswana': '茨瓦纳语', 'tn': '茨瓦纳语', 'afrikaans': '南非荷兰语', 'afr': '南非荷兰语',
    'ndebele': '恩德贝勒语', 'xhosa': '科萨语', 'xho': '科萨语',
    'assamese': '阿萨姆语', 'oriya)': '奥里亚语',
    'saraiki': '萨拉基语', 'braj': '布拉杰语', 'bhojpuri': '博杰普尔语',
    'kumaoni': '库马奥尼语', 'garhwali': '加瓦尔语', 'dogri': '多格拉语',
    'rajasthani': '拉贾斯坦语', 'marwari': '马尔瓦里语',
    'tulu': '图鲁语', 'konkani': '孔卡尼语', ' manipuri': '曼尼普尔语',
    'bodo': '博多语', 'khasi': '卡西语', 'garo': '加罗语', 'mizo': '米佐语',
    'lushai': '米佐语', 'naga': '那加语', 'kuki': '库基语', 'chin': '钦语',
    'jinghpaw': '景颇语', 'jingpho': '景颇语', 'kachin': '景颇语',
    'wa': '佤语', 'va': '佤语', 'palaung': '德昂语', 'blang': '布朗语',
    'achang': '阿昌语', 'nuosu': '彝语', 'yi': '彝语',
    'karén': '克伦语', 'karenni': '克耶语', 'kayah': '克耶语',
    'mon': '蒙语', 'mongolian)': '蒙语', 'burmese ethnic': '缅甸民族语言',
    'sgaw': '克伦语', 'pwo': '克伦语',
    # ---- 青木表多词语言名（HFCC ISO 639-3 英文名） ----
    'mandarin chinese': '中文', 'standard chinese': '中文', 'standart chinese': '中文',
    'min nan chinese': '闽南语', 'hakka chinese': '客家话', 'yue chinese': '粤语',
    'multiple languages': '多语种', 'not identified': '未指定', 'undetermined': '未指定',
    'standard arabic': '阿拉伯语', 'sudanese arabic': '阿拉伯语', 'gulf arabic': '阿拉伯语',
    'mesopotamian arabic': '阿拉伯语', 'north levantine arabic': '阿拉伯语',
    'central pashto': '普什图语', 'northern pashto': '普什图语', 'southern pashto': '普什图语',
    'central khmer': '高棉语', 'kirghiz': '吉尔吉斯语', 'fulah': '富拉语',
    'macedo romanian': '阿罗马尼亚语', 'sidamo': '锡达莫语',
    "s'gaw karen": '克伦语', 'manumanaw karen': '克伦语', 'pwo karen': '克伦语',
    'middle armenian': '亚美尼亚语', 'tigre': '提格雷语',
    'rumai palaung': '德昂语', 'palaung': '德昂语', 'khamti': '掸语',
    'hmong daw': '苗语', 'hmong njua': '苗语', 'yao': '瑶语', 'makhuwa': '马库阿语',
    'swiss german': '德语', 'uighur': '维吾尔语', 'kurmanji': '库尔德语',
    'sorani': '库尔德语', 'plesk': '未指定',
}


# 语言分组英文名（供英文界面使用）
GROUP_EN = {
    '中文': 'Mandarin Chinese', '粤语': 'Cantonese', '英语': 'English', '日语': 'Japanese',
    '韩语': 'Korean', '俄语': 'Russian', '阿拉伯语': 'Arabic', '西班牙语': 'Spanish',
    '法语': 'French', '德语': 'German', '葡语': 'Portuguese', '意语': 'Italian',
    '印地语': 'Hindi', '乌尔都语': 'Urdu', '孟加拉语': 'Bengali', '泰米尔语': 'Tamil',
    '泰卢固语': 'Telugu', '马拉雅拉姆语': 'Malayalam', '维吾尔语': 'Uyghur', '藏语': 'Tibetan',
    '蒙语': 'Mongolian', '越南语': 'Vietnamese', '泰语': 'Thai', '缅甸语': 'Burmese',
    '印尼语': 'Indonesian', '马来语': 'Malay', '波斯语': 'Persian', '达里语': 'Dari',
    '普什图语': 'Pashto', '土耳其语': 'Turkish', '波兰语': 'Polish', '罗马尼亚语': 'Romanian',
    '保加利亚语': 'Bulgarian', '乌克兰语': 'Ukrainian', '斯瓦希里语': 'Swahili', '豪萨语': 'Hausa',
    '索马里语': 'Somali', '阿姆哈拉语': 'Amharic', '克林冈语': 'Klingon', '世界语': 'Esperanto',
    '多语种': 'Multiple languages', '未指定': 'Unspecified', '授时信号': 'Time signal',
    '干扰信号': 'Jamming', '数字数据': 'Data', '音乐': 'Music', '客家话': 'Hakka',
    '闽南语': 'Hokkien', '潮州话': 'Teochew', '吴语': 'Wu', '粤语（广州话）': 'Cantonese',
    '当地语言': 'Vernacular', '库尔德语': 'Kurdish', '希腊语': 'Greek', '荷兰语': 'Dutch',
    '高棉语': 'Khmer', '老挝语': 'Lao', '僧伽罗语': 'Sinhala', '尼泊尔语': 'Nepali',
    '斯洛伐克语': 'Slovak', '德语（瑞士）': 'German', '富拉语': 'Fulah', '阿罗马尼亚语': 'Aromanian',
    '克伦语': 'Karen', '亚美尼亚语': 'Armenian', '提格雷尼亚语': 'Tigrinya', '苗语': 'Hmong',
    '傈僳语': 'Lisu', '拉祜语': 'Lahu', '掸语': 'Shan', '若开语': 'Rakhine', '罗兴亚语': 'Rohingya',
    '卡纳达语': 'Kannada', '班巴拉语': 'Bambara', '拉丁语': 'Latin', '沃洛夫语': 'Wolof',
    '契维语': 'Twi', '埃维语': 'Ewe', '苏苏语': 'Susu', '约鲁巴语': 'Yoruba', '伊博语': 'Igbo',
    '林加拉语': 'Lingala', '桑戈语': 'Sango', '聪加语': 'Tsonga', '茨瓦纳语': 'Tswana',
    '南非荷兰语': 'Afrikaans', '恩德贝勒语': 'Ndebele', '科萨语': 'Xhosa', '祖鲁语': 'Zulu',
    '阿法尔语': 'Afar', '奥罗莫语': 'Oromo', '锡达莫语': 'Sidamo', '马库阿语': 'Makhuwa',
    '克丘亚语': 'Quechua', '艾马拉语': 'Aymara', '瓜拉尼语': 'Guarani', '克里奥尔语': 'Creole',
    '德昂语': 'Palaung', '瑶语': 'Yao', '景颇语': 'Jinghpaw', '佤语': 'Wa',
    '比斯拉马语': 'Bislama', '斐济语': 'Fijian', '萨摩亚语': 'Samoan', '巴布亚皮钦语': 'Tok Pisin',
    '希里莫图语': 'Hiri Motu', '毛利语': 'Maori', '琉森语': 'Rundi',
    '白俄罗斯语': 'Belarusian', '塞尔维亚语': 'Serbian', '克罗地亚语': 'Croatian',
    '斯洛文尼亚语': 'Slovenian', '马其顿语': 'Macedonian', '阿尔巴尼亚语': 'Albanian',
    '匈牙利语': 'Hungarian', '捷克语': 'Czech', '立陶宛语': 'Lithuanian', '拉脱维亚语': 'Latvian',
    '爱沙尼亚语': 'Estonian', '芬兰语': 'Finnish', '瑞典语': 'Swedish', '挪威语': 'Norwegian',
    '丹麦语': 'Danish', '冰岛语': 'Icelandic', '希伯来语': 'Hebrew', '格鲁吉亚语': 'Georgian',
    '阿塞拜疆语': 'Azerbaijani', '哈萨克语': 'Kazakh', '乌兹别克语': 'Uzbek', '土库曼语': 'Turkmen',
    '塔吉克语': 'Tajik', '吉尔吉斯语': 'Kyrgyz', '信德语': 'Sindhi', '旁遮普语': 'Punjabi',
    '古吉拉特语': 'Gujarati', '马拉地语': 'Marathi', '奥里亚语': 'Oriya', '阿萨姆语': 'Assamese',
    '尼泊尔语（不丹）': 'Nepali', '俾路支语': 'Balochi', '克什米尔语': 'Kashmiri',
    '缅甸民族语言': 'Burmese ethnic', '克耶语': 'Kayah', '德顿语': 'Tetum',
}

def lang_bucket(raw):
    """原始语言描述 -> (显示原文, 分组名)"""
    if raw is None:
        return '', '未指定'
    s = str(raw).strip()
    if not s or s.lower() in ('none', 'null'):
        return '', '未指定'
    first = re.split(r'[/,;()\[\]·"\']', s)[0].strip().lower().strip(' .')
    return s, LANG_MAP.get(first, s if first else '未指定')

def band_of(f):
    """kHz -> 频段键"""
    if f < 148:
        return 'time'          # 标准时间/频率信号台
    if f < 526.5:
        return 'LW'
    if f < 2300:
        return 'MW'
    if f < 87500:
        return 'SW'
    if f <= 108000:
        return 'FM'
    return 'other'

def parse_utc_window(s):
    """'HHMM-HHMM' -> (start_min, end_min)；'0000-2400' -> (0,1440)；异常返回 None"""
    s = str(s).strip()
    m = re.fullmatch(r'(\d{2})(\d{2})-(\d{2})(\d{2})', s)
    if not m:
        return None
    a = int(m.group(1)) * 60 + int(m.group(2))
    b = int(m.group(3)) * 60 + int(m.group(4))
    if a == 0 and b == 0:   # '0000-0000' 视为全天
        b = 1440
    return a, b

def norm_days(s):
    s = str(s or '').replace('.', '').replace(' ', '')
    s = ''.join(c for c in s if c in '1234567')
    return s or '1234567'

def clean(v, maxlen=90):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v).strip())[:maxlen]

def flag(iso2):
    if not iso2:
        return ''
    return ''.join(chr(0x1F1E6 + ord(c) - 65) for c in iso2.upper())

# ---------------- xta26 表 ----------------
def parse_xta26(ws):
    out = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2:
            continue
        if row[0] is None:
            continue
        try:
            f = float(row[0])
        except (TypeError, ValueError):
            continue
        win = parse_utc_window(row[3] or '')
        if win is None:
            continue
        s, e = win
        station = clean(row[2], 70)
        lang_disp, lang_n = lang_bucket(row[6])
        joined = ' '.join(str(c) for c in row if c is not None).upper()
        drm = 1 if ('DRM' in joined or '(DIGITAL' in joined or 'DIGITAL)' in joined) else 0
        cc = str(row[10] or '').strip()
        cn, en, iso = COUNTRIES.get(cc, (cc, cc, ''))
        pw = clean(row[7], 12)
        ll = None
        mll = re.match(r'^(\d{2})(\d{2})(\d{2})([NS])(\d{3})(\d{2})(\d{2})([EW])$', clean(row[11], 20))
        if mll:
            ll = [_dm_to_deg(mll.group(1), mll.group(2), mll.group(4)),
                  _dm_to_deg(mll.group(5), mll.group(6), mll.group(8))]
        out.append({
            'f': f, 's': s, 'e': e, 'd': norm_days(row[5]),
            'st': station, 'lg': lang_disp, 'lgn': lang_n,
            'pw': pw, 'az': clean(row[8], 10), 'si': clean(row[9], 45),
            'cc': cc, 'cn': cn, 'iso': iso, 'drm': drm,
            'src': 'xta26', 'nt': clean(row[12], 70),
        })
        if ll:
            out[-1]['ll'] = ll
    return out

# ---------------- HFCC 标准格式 TXT（B26 等，直接来自 hfcc.org 数据包） ----------------
def _dm_to_deg(d, m, hemi):
    v = int(d) + int(m) / 60.0
    return round(-v if hemi in 'SW' else v, 2)

def load_ref_tables(refdir):
    """site/broadcas/admin 参考表 → 查询 dict（site 含经纬度 [lat, lon]）"""
    sites, brcs, adms = {}, {}, {}
    def rows(name):
        return [l.rstrip() for l in open(os.path.join(refdir, name), encoding='utf-8-sig', errors='replace')
                if l.strip() and not l.lstrip().startswith(';')]
    try:
        for l in rows('site.txt'):
            code = l[:3].strip()
            if not code:
                continue
            ll = None
            # site.txt 坐标格式: 43N17 077E00（度 半球 分）
            m = re.search(r'(\d{1,3})([NS])(\d{2})\s+(\d{1,3})([EW])(\d{2})\s*$', l)
            if m:
                ll = [_dm_to_deg(m.group(1), m.group(3), m.group(2)),
                      _dm_to_deg(m.group(4), m.group(6), m.group(5))]
            sites[code] = (l[4:35].strip() or code, l[35:39].strip(), ll)
    except OSError:
        pass
    try:
        for l in rows('broadcas.txt'):
            code = l[:3].strip()
            if code:
                brcs[code] = l[4:].strip() or code
    except OSError:
        pass
    try:
        for l in rows('admin.txt'):
            code = l[:3].strip()
            if code:
                adms[code] = l[4:55].strip() or code
    except OSError:
        pass
    return sites, brcs, adms

def parse_hfcc_txt(path, refdir=None):
    """解析 HFCC 标准格式季表（B26all00.TXT），固定列宽按标尺切分"""
    lines = open(path, encoding='utf-8', errors='replace').read().splitlines()
    ruler = next(l for l in lines if l.startswith(';--'))
    plus = [i for i, c in enumerate(ruler) if c == '+']
    starts = [0] + [p + 1 for p in plus]
    ends = plus + [9999]
    names = ['FREQ', 'STRT', 'STOP', 'CIRAF', 'LOC', 'POWR', 'AZIMUTH', 'SLW', 'ANT', 'DAYS',
             'FDATE', 'TDATE', 'MOD', 'AFRQ', 'LANGUAGE', 'ADM', 'BRC', 'FMO', 'REQ',
             'OLD', 'ALT1', 'ALT2', 'ALT3', 'NOTES']
    sites, brcs, adms = load_ref_tables(refdir) if refdir else ({}, {}, {})
    today = __import__('datetime').date.today()

    def ddmmmy(s):
        return datetime.date(2000 + int(s[4:]), int(s[2:4]), int(s[:2]))

    out = []
    for l in lines:
        if not l.strip() or l.lstrip().startswith(';'):
            continue
        f = {n: l[starts[i]:ends[i]].strip() for i, n in enumerate(names)}
        if not f['FREQ'].isdigit():
            continue
        # 季节有效性：FDATE ≤ 今天 ≤ TDATE
        try:
            if f['FDATE'] and ddmmmy(f['FDATE']) > today:
                continue
            if f['TDATE'] and ddmmmy(f['TDATE']) < today:
                continue
        except ValueError:
            pass
        s, e = int(f['STRT'][:2]) * 60 + int(f['STRT'][2:4]), int(f['STOP'][:2]) * 60 + int(f['STOP'][2:4])
        if s == 0 and e == 0:
            e = 1440
        lang_disp, lang_n = lang_bucket(f['LANGUAGE'])
        adm = f['ADM']
        if adm in COUNTRIES:
            cn, _en, iso = COUNTRIES[adm]
        else:
            cn, iso = adms.get(adm, adm or ''), ''
        _site = sites.get(f['LOC'], (f['LOC'], '', None))
        site_name = _site[0]
        brc_name = brcs.get(f['BRC'], f['BRC'])
        out.append({
            'f': int(f['FREQ']), 's': s, 'e': e, 'd': norm_days(f['DAYS']),
            'st': clean(brc_name, 48), 'lg': lang_disp, 'lgn': lang_n,
            'pw': clean(f['POWR'], 12), 'az': clean(f['AZIMUTH'], 6), 'si': clean(site_name, 40),
            'cc': adm, 'cn': cn, 'iso': iso, 'drm': 1 if f['MOD'] == 'N' else 0,
            'src': 'hfcc', 'fmo': clean(f['FMO'], 24), 'tgt': clean(f['CIRAF'], 40),
        })
        if _site[2]:
            out[-1]['ll'] = _site[2]
    return out


# ---------------- 青木格式频率表（用户提供 xlsx：主表 + HFCC 原始表） ----------------
def _aoki_hdr(rows, keys):
    """找到包含全部关键列名的表头行，返回 列名->列号 映射"""
    for i, r in enumerate(rows[:10]):
        names = {str(c).strip(): j for j, c in enumerate(r) if c is not None}
        if all(k in names for k in keys):
            return names
    raise ValueError('未找到表头行（需要列: %s）' % ','.join(keys))

def parse_aoki(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    keys = ['FRE', 'STATION', 'UTC', 'Su-W-Sa', 'Language', 'Pow', 'Azi',
            'Location', 'ADM', 'L/L', 'Remarks', 'CIRAF Zones', 'REQ#', 'FDATE', 'TDATE']
    h = _aoki_hdr(rows, keys)
    g = lambda r, k: (r[h[k]] if h[k] < len(r) else None)

    # HFCC 原始表：REQ# -> MOD（N = DRM）
    req_mod = {}
    try:
        ws2 = wb['HFCC原始数据']
        rows2 = list(ws2.iter_rows(values_only=True))
        h2 = _aoki_hdr(rows2, ['REQ#', 'MOD'])
        for r in rows2:
            if r[h2['REQ#']] is None:
                continue
            req_mod[str(r[h2['REQ#']])] = str(r[h2['MOD']])
    except (KeyError, ValueError):
        pass

    today = datetime.date.today()

    def ddmmmy(s):
        return datetime.date(2000 + int(s[4:]), int(s[2:4]), int(s[:2]))

    out = []
    for r in rows:
        if r[h['FRE']] is None:
            continue
        try:
            f = int(float(g(r, 'FRE')))
        except (TypeError, ValueError):
            continue
        win = parse_utc_window(g(r, 'UTC') or '')
        if win is None:
            continue
        s_, e_ = win
        # 季节有效性：FDATE ≤ 今天 ≤ TDATE
        try:
            fd, td = clean(g(r, 'FDATE'), 8), clean(g(r, 'TDATE'), 8)
            if fd and ddmmmy(fd) > today:
                continue
            if td and ddmmmy(td) < today:
                continue
        except ValueError:
            pass
        lang_disp, lang_n = lang_bucket(g(r, 'Language'))
        cc = str(g(r, 'ADM') or '').strip()
        cn, _en, iso = COUNTRIES.get(cc, (cc, cc, ''))
        ce = COUNTRIES.get(cc, ('', cc, ''))[1] or cc
        ll = None
        mll = re.match(r'^(\d{2})(\d{2})(\d{2})([NS])(\d{3})(\d{2})(\d{2})([EW])$', clean(g(r, 'L/L'), 20))
        if mll:
            ll = [_dm_to_deg(mll.group(1), mll.group(2), mll.group(4)),
                  _dm_to_deg(mll.group(5), mll.group(6), mll.group(8))]
        req = str(g(r, 'REQ#') or '').strip()
        it = {
            'f': f, 's': s_, 'e': e_, 'd': norm_days(g(r, 'Su-W-Sa')),
            'st': clean(g(r, 'STATION'), 60), 'lg': lang_disp, 'lgn': lang_n,
            'lge': GROUP_EN.get(lang_n, lang_disp or lang_n or 'Unspecified'),
            'ce': ce,
            'pw': clean(g(r, 'Pow'), 12), 'az': clean(g(r, 'Azi'), 10),
            'si': clean(g(r, 'Location'), 45), 'cc': cc, 'cn': cn, 'iso': iso,
            'drm': 1 if req_mod.get(req) == 'N' else 0,
            'src': 'aoki', 'nt': clean(g(r, 'Remarks'), 70), 'tgt': clean(g(r, 'CIRAF Zones'), 40),
        }
        if ll:
            it['ll'] = ll
        out.append(it)
    return out

# ---------------- FM：radiomap.eu 城市台 ----------------
# slug -> (城市标签, IANA 时区, ISO2 国码, 中文国名, 英文国名)
FM_CITIES = {
    'al_tirana': ('地拉那 Tirana', 'Europe/Tirane', 'AL', '阿尔巴尼亚'),
    'at_wien': ('维也纳 Vienna', 'Europe/Vienna', 'AT', '奥地利'),
    'ba_sarajevo': ('萨拉热窝 Sarajevo', 'Europe/Sarajevo', 'BA', '波黑'),
    'be_bruxelles': ('布鲁塞尔 Brussels', 'Europe/Brussels', 'BE', '比利时'),
    'ch_bern': ('伯尔尼 Bern', 'Europe/Zurich', 'CH', '瑞士'),
    'ch_geneve': ('日内瓦 Geneva', 'Europe/Zurich', 'CH', '瑞士'),
    'ch_zuerich': ('苏黎世 Zurich', 'Europe/Zurich', 'CH', '瑞士'),
    'cy_lefkosia': ('尼科西亚 Nicosia', 'Asia/Nicosia', 'CY', '塞浦路斯'),
    'cz_praha': ('布拉格 Prague', 'Europe/Prague', 'CZ', '捷克'),
    'de_berlin': ('柏林 Berlin', 'Europe/Berlin', 'DE', '德国'),
    'de_hamburg': ('汉堡 Hamburg', 'Europe/Berlin', 'DE', '德国'),
    'de_koeln': ('科隆 Cologne', 'Europe/Berlin', 'DE', '德国'),
    'de_muenchen': ('慕尼黑 Munich', 'Europe/Berlin', 'DE', '德国'),
    'dk_kobenhavn': ('哥本哈根 Copenhagen', 'Europe/Copenhagen', 'DK', '丹麦'),
    'ee_tallinn': ('塔林 Tallinn', 'Europe/Tallinn', 'EE', '爱沙尼亚'),
    'es_barcelona': ('巴塞罗那 Barcelona', 'Europe/Madrid', 'ES', '西班牙'),
    'es_madrid': ('马德里 Madrid', 'Europe/Madrid', 'ES', '西班牙'),
    'fi_helsinki': ('赫尔辛基 Helsinki', 'Europe/Helsinki', 'FI', '芬兰'),
    'fr_marseille': ('马赛 Marseille', 'Europe/Paris', 'FR', '法国'),
    'fr_paris': ('巴黎 Paris', 'Europe/Paris', 'FR', '法国'),
    'gr_athina': ('雅典 Athens', 'Europe/Athens', 'GR', '希腊'),
    'hr_zagreb': ('萨格勒布 Zagreb', 'Europe/Zagreb', 'HR', '克罗地亚'),
    'hu_budapest': ('布达佩斯 Budapest', 'Europe/Budapest', 'HU', '匈牙利'),
    'ie_dublin': ('都柏林 Dublin', 'Europe/Dublin', 'IE', '爱尔兰'),
    'is_reykjavik': ('雷克雅未克 Reykjavík', 'Atlantic/Reykjavik', 'IS', '冰岛'),
    'it_milano': ('米兰 Milan', 'Europe/Rome', 'IT', '意大利'),
    'it_roma': ('罗马 Rome', 'Europe/Rome', 'IT', '意大利'),
    'lt_vilnius': ('维尔纽斯 Vilnius', 'Europe/Vilnius', 'LT', '立陶宛'),
    'lu_letzebuerg': ('卢森堡城 Luxembourg City', 'Europe/Luxembourg', 'LU', '卢森堡'),
    'lv_riga': ('里加 Riga', 'Europe/Riga', 'LV', '拉脱维亚'),
    'mc_monaco': ('摩纳哥 Monaco', 'Europe/Monaco', 'MC', '摩纳哥'),
    'md_chisinau': ('基希讷乌 Chișinău', 'Europe/Chisinau', 'MD', '摩尔多瓦'),
    'mk_skopje': ('斯科普里 Skopje', 'Europe/Skopje', 'MK', '北马其顿'),
    'mt_valletta': ('瓦莱塔 Valletta', 'Europe/Malta', 'MT', '马耳他'),
    'nl_amsterdam': ('阿姆斯特丹 Amsterdam', 'Europe/Amsterdam', 'NL', '荷兰'),
    'no_oslo': ('奥斯陆 Oslo', 'Europe/Oslo', 'NO', '挪威'),
    'pl_warszawa': ('华沙 Warsaw', 'Europe/Warsaw', 'PL', '波兰'),
    'pt_lisboa': ('里斯本 Lisbon', 'Europe/Lisbon', 'PT', '葡萄牙'),
    'ro_bucuresti': ('布加勒斯特 Bucharest', 'Europe/Bucharest', 'RO', '罗马尼亚'),
    'rs_beograd': ('贝尔格莱德 Belgrade', 'Europe/Belgrade', 'RS', '塞尔维亚'),
    'se_stockholm': ('斯德哥尔摩 Stockholm', 'Europe/Stockholm', 'SE', '瑞典'),
    'si_ljubljana': ('卢布尔雅那 Ljubljana', 'Europe/Ljubljana', 'SI', '斯洛文尼亚'),
    'sk_bratislava': ('布拉迪斯拉发 Bratislava', 'Europe/Bratislava', 'SK', '斯洛伐克'),
    'tr_ankara': ('安卡拉 Ankara', 'Europe/Istanbul', 'TR', '土耳其'),
    'tr_istanbul': ('伊斯坦布尔 Istanbul', 'Europe/Istanbul', 'TR', '土耳其'),
    'ua_kyiv': ('基辅 Kyiv', 'Europe/Kyiv', 'UA', '乌克兰'),
    'uk_aberdeen': ('阿伯丁 Aberdeen', 'Europe/London', 'GB', '英国'),
    'uk_belfast': ('贝尔法斯特 Belfast', 'Europe/London', 'GB', '英国'),
    'uk_birmingham': ('伯明翰 Birmingham', 'Europe/London', 'GB', '英国'),
    'uk_bristol': ('布里斯托尔 Bristol', 'Europe/London', 'GB', '英国'),
    'uk_cardiff': ('加的夫 Cardiff', 'Europe/London', 'GB', '英国'),
    'uk_edinburgh': ('爱丁堡 Edinburgh', 'Europe/London', 'GB', '英国'),
    'uk_london': ('伦敦 London', 'Europe/London', 'GB', '英国'),
    'uk_manchester': ('曼彻斯特 Manchester', 'Europe/London', 'GB', '英国'),
}

ISO2_EN = {
    'AL': 'Albania', 'AT': 'Austria', 'BA': 'Bosnia & Herzegovina', 'BE': 'Belgium',
    'CH': 'Switzerland', 'CY': 'Cyprus', 'CZ': 'Czechia', 'DE': 'Germany', 'DK': 'Denmark',
    'EE': 'Estonia', 'ES': 'Spain', 'FI': 'Finland', 'FR': 'France', 'GB': 'United Kingdom',
    'GR': 'Greece', 'HR': 'Croatia', 'HU': 'Hungary', 'IE': 'Ireland', 'IS': 'Iceland',
    'IT': 'Italy', 'LT': 'Lithuania', 'LU': 'Luxembourg', 'LV': 'Latvia', 'MC': 'Monaco',
    'MD': 'Moldova', 'MK': 'North Macedonia', 'MT': 'Malta', 'NL': 'Netherlands',
    'NO': 'Norway', 'PL': 'Poland', 'PT': 'Portugal', 'RO': 'Romania', 'RS': 'Serbia',
    'SE': 'Sweden', 'SI': 'Slovenia', 'SK': 'Slovakia', 'TR': 'Turkey', 'UA': 'Ukraine',
}

def parse_radiomap(pagedir):
    """解析 radiomap.eu 城市页（<cc>_<city>.htm）→ FM 条目（全天播出）"""
    import re as _re
    out = []
    for fn in sorted(os.listdir(pagedir)):
        if not fn.endswith('.htm'):
            continue
        slug = fn[:-4]
        if slug not in FM_CITIES:
            continue
        label, tz, iso, country = FM_CITIES[slug]
        city_en = label.split(' ', 1)[1] if ' ' in label else label
        html = open(os.path.join(pagedir, fn), encoding='utf-8', errors='replace').read()
        for rowm in _re.finditer(r'<tr[^>]*>(.*?)</tr>', html, _re.S):
            cells = _re.findall(r'<td[^>]*>(.*?)</td>', rowm.group(1), _re.S)
            if len(cells) < 2:
                continue
            fm = _re.search(r'(\d{2,3}(?:\.\d)?)\s*$', re.sub(r'<[^>]+>', '', cells[0]).replace('&#160;', ' ').strip())
            if not fm:
                continue
            try:
                f = float(fm.group(1))
            except ValueError:
                continue
            if not 76 <= f <= 108:          # 合法 FM 频段 (MHz)
                continue
            parts = [_re.sub(r'\s+', ' ', _re.sub(r'<[^>]+>', ' ', c)).strip() for c in cells[1].split('<br')]
            name = ' / '.join(p.lstrip('> ').strip() for p in parts if p.lstrip('> ').strip())[:90] or '未知电台'
            site = _re.sub(r'\s+', ' ', _re.sub(r'<[^>]+>', ' ', cells[2])).strip() if len(cells) > 2 else ''
            out.append({
                'f': round(f * 1000, 1), 's': 0, 'e': 1440, 'd': '1234567',
                'st': name, 'lg': '', 'lgn': '未指定', 'lge': 'Unspecified',
                'si': clean(site, 40), 'cc': iso, 'cn': country, 'ce': ISO2_EN.get(iso, iso), 'iso': iso,
                'src': 'fm', 'tz': tz, 'city': label, 'cty': city_en, 'drm': 0,
            })
    return out

# ---------------- hfcc0904 表（A26 xlsx 版，留作备用；现用 B26 TXT） ----------------
def parse_hfcc(ws):
    out = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2:
            continue
        if row[0] is None:
            continue
        try:
            f = int(float(row[0]))
        except (TypeError, ValueError):
            continue
        try:
            a = str(row[1]).strip(); b = str(row[2]).strip()
            s = int(a[:2]) * 60 + int(a[2:4])
            e = int(b[:2]) * 60 + int(b[2:4])
            if b == '2400':
                e = 1440
            if s == 0 and e == 0:
                e = 1440
        except (TypeError, ValueError):
            continue
        lang_disp, lang_n = lang_bucket(row[16])
        cc = str(row[17] or '').strip()
        cn, en, iso = COUNTRIES.get(cc, (cc, cc, ''))
        out.append({
            'f': f, 's': s, 'e': e, 'd': norm_days(row[11]),
            'st': clean(row[18], 40), 'lg': lang_disp, 'lgn': lang_n,
            'pw': clean(row[7], 12), 'az': clean(row[8], 6), 'si': clean(row[6], 24),
            'cc': cc, 'cn': cn, 'iso': iso, 'drm': 1 if str(row[14]) == 'N' else 0,
            'src': 'hfcc', 'brc': clean(row[18], 40), 'fmo': clean(row[19], 24),
            'tgt': clean(row[5], 40),
        })
    return out

# ---------------- WRMI 网格 (rows 4-99, 4 行/小时块) ----------------
def parse_wrmi(ws):
    rows = list(ws.iter_rows(values_only=True))
    tx_freq = {}
    tx_az = {}
    for c in range(2, 15):
        cell = re.sub(r'\D', '', str(rows[2][c] or ''))
        if cell:
            tx_freq[c] = int(cell)
            tx_az[c] = clean(rows[3][c], 24)
    out = []
    h = 4
    while h + 3 < len(rows) and h <= 99:
        hourcell = rows[h][0]
        m = re.fullmatch(r'\s*(\d{4})\s*', str(hourcell or ''))
        if not m:
            h += 1
            continue
        hh = int(m.group(1)[:2]); mm = int(m.group(1)[2:])
        start = hh * 60 + mm
        end = start + 60
        for c, fq in tx_freq.items():
            parts = []
            for r in (h + 1, h + 2, h + 3):
                try:
                    v = rows[r][c]
                except IndexError:
                    v = None
                if v is not None:
                    parts.append(str(v).strip())
            name = re.sub(r'\s+', ' ', ' '.join(p for p in parts if p)).strip()
            letter = str(rows[h][c] or '').strip()
            lg_disp = ''
            lgn = '未指定'
            if not name:
                if not re.fullmatch(r'[A-N]', letter):
                    continue                      # 台关机：无字母无名字
                name = 'WRMI 节目（代码 %s，名称见图例）' % letter
            else:
                low = name.lower()
                for kw, b in (('français', '法语'), ('español', '西班牙语'),
                              ('portugu', '葡语'), ('deutsch', '德语'),
                              ('slovak', '斯洛伐克语'), ('slovaque', '斯洛伐克语'),
                              ('english', '英语'), ('esl', '西班牙语')):
                    if kw in low:
                        lgn = b
                        break
            out.append({
                'f': fq, 's': start, 'e': end % 1440 if end % 1440 else 1440,
                'e_abs': end, 'd': '1234567',
                'st': 'WRMI', 'nm': name[:110], 'lg': lg_disp, 'lgn': lgn,
                'pw': '100', 'az': tx_az.get(c, ''), 'si': 'Okeechobee FL',
                'cc': 'USA', 'cn': '美国', 'iso': 'US', 'drm': 0,
                'src': 'wrmi', 'll': [27.53, -80.91],
            })
        h += 4
    return out

# ---------------- 合并 & 去重 & DRM 交叉富化 ----------------
def merge(xta, hfcc, wrmi):
    # 1) 去重：key=(f,s,e,lgn) 优先级 wrmi > xta26 > hfcc
    prio = {'wrmi': 0, 'xta26': 1, 'hfcc': 2}
    best = {}
    for lst, src in ((wrmi, 'wrmi'), (xta, 'xta26'), (hfcc, 'hfcc')):
        for it in lst:
            k = (it['f'], it['s'], it['e'], it['lgn'], it.get('tz', ''))
            cur = best.get(k)
            if cur is None or prio[src] < prio[cur['src']]:
                best[k] = it
    merged = list(best.values())

    # 2) DRM 富化：hfcc MOD=N 的 (freq,时段) 标记到同频同时段的非 DRM 条目
    drm_idx = collections.defaultdict(list)
    for it in merged:
        if it['drm'] and it['src'] == 'hfcc':
            drm_idx[it['f']].append((it['s'], it['e']))
    for it in merged:
        if it['drm']:
            continue
        for (s2, e2) in drm_idx.get(it['f'], ()):
            # 时段重叠即视为同一数字发射
            if not (e2 <= it['s'] or s2 >= it['e']):
                it['drm'] = 1
                break
    return merged

def main():
    """用法:
      python tools/build_data.py --schedule <青木格式xlsx> [--out data] [--radiomap <目录>]
      --schedule : HFCC 季表（青木格式）xlsx，主表 + HFCC原始数据 表（含 MOD 判 DRM）
      --radiomap : radiomap.eu 城市页目录（<cc>_<city>.htm），提供时生成 FM 条目
    """
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('--schedule', required=True, help='青木格式频率表 xlsx')
    ap.add_argument('--out', default=os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data'))
    ap.add_argument('--radiomap', default=None, help='radiomap.eu 城市页目录')
    args = ap.parse_args()
    outdir = args.out
    os.makedirs(outdir, exist_ok=True)

    sched = parse_aoki(args.schedule)
    fm = parse_radiomap(args.radiomap) if args.radiomap and os.path.isdir(args.radiomap) else []

    merged = sched + fm
    for it in merged:
        it['band'] = band_of(it['f'])
        # 输出瘦身：丢空字段
        it.pop('brc', None)
        for k in [k for k, v in it.items() if v in ('', None, 0) and k not in ('f', 's', 'e', 'd', 'drm')]:
            it.pop(k, None)

    data = {
        'meta': {
            'title': 'HFCC B26 频率表（青木格式） / radiomap.eu FM',
            'season': 'B26',
            'updated': '2026-10-30',
            'generated': datetime.date.today().isoformat(),
            'sources': {
                'aoki': 'HFCC B26 青木格式频率表（短波 2485~26040 kHz，DRM=原始表 MOD"N"，已按 FDATE/TDATE 过滤）',
                'fm': 'radiomap.eu 各城市 FM 电台频率（全天播出）',
            },
        },
        'entries': merged,
    }
    out = os.path.join(outdir, 'data.js')
    with open(out, 'w', encoding='utf-8') as fp:
        fp.write('/* 由 tools/build_data.py 自动生成 — 请勿手工编辑 */\n')
        fp.write('window.RADIO_DATA=')
        fp.write(json.dumps(data, ensure_ascii=False, separators=(',', ':')))
        fp.write(';\n')

    # 统计
    bands = collections.Counter(it['band'] for it in merged)
    drm_n = sum(it['drm'] for it in merged)
    print(f'aoki: {len(sched)}  fm: {len(fm)}  merged: {len(merged)}')
    print('bands:', dict(bands), ' DRM:', drm_n)
    langs = collections.Counter(it['lgn'] for it in merged)
    print('top langs:', langs.most_common(12))
    print('written:', os.path.abspath(out), f'{os.path.getsize(out)/1e6:.1f} MB')

if __name__ == '__main__':
    main()
