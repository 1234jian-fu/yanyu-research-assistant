"""
瀛︾爺路宸ョ绉戠爺鍔╂墜 v4.0 (YanYu OS)
15+鍔熻兘鐭╅樀 路 鏉垮潡閿氬畾 路 闆剁浣嶆墽琛?路 褰卞瓙鍚堣憲鑰?路 璇█鍩哄洜娣卞害鎻愬彇
"""

import os
import re
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from collections import Counter
import json

import anthropic
import fitz  # pymupdf
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.shared import OxmlElement
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# 鈹€鈹€ Configuration 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€
st.set_page_config(
    page_title="瀛︾爺路宸ョ绉戠爺鍔╂墜 v4.0",
    layout="wide",
    initial_sidebar_state="expanded",
)

<<<<<<< HEAD
# Claude API 閰嶇疆 (浣跨敤 Claude Code 鐩稿悓鎺ュ彛)
CLAUDE_API_KEY = (os.getenv("ANTHROPIC_API_KEY") or os.getenv("ANTHROPIC_AUTH_TOKEN") or "").strip()
CLAUDE_BASE_URL = os.getenv("ANTHROPIC_BASE_URL", "https://api.anthropic.com").rstrip("/")
CLAUDE_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-6")
CLAUDE_API_KEY = os.getenv("ANTHROPIC_AUTH_TOKEN", "").strip()
CLAUDE_BASE_URL = os.getenv("ANTHROPIC_BASE_URL", "https://aiapi.aixia.tech").rstrip("/")
CLAUDE_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-opus-4-6")
CLAUDE_MODEL_OPTIONS = [
    "claude-opus-4-6",
    "claude-sonnet-4-6",
]

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_BASE_URL = os.getenv("GEMINI_BASE_URL", "https://new.lemonapi.site").rstrip("/")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "imagen-2.0-generate-001")
GEMINI_MODEL_OPTIONS = [
    "imagen-2.0-generate-001",
    "gemini-3.1-flash-image-preview",
    "gemini-2.5-flash-image-preview",
    "gemini-2.0-flash-preview-image-generation",
]

WRITING_PAGE = "鉁嶏笍 璁烘枃鍐欎綔"
FORMATTING_PAGE = "馃搻 鏍煎紡瀵归綈"
VIZ_PAGE = "馃帹 瑙嗚瀹為獙瀹?
PPT_PAGE = "馃梻锔?PPT澶у笀"

CONFIG = {
    "sections": {
        "鎽樿": {
            "focus": "寮€闂ㄨ灞憋紝鏁版嵁鏀拺",
            "max_words": 250,
            "goal": "鐢ㄦ渶鐭瘒骞呬氦浠ｇ爺绌跺璞°€佹柟娉曘€佹牳蹇冪粨鏋滀笌璐＄尞銆?,
            "must_include": ["鐮旂┒瀵硅薄/闂", "閲囩敤鐨勬柟娉曟垨绛栫暐", "鍏抽敭缁撴灉鎴栨暟鎹秼鍔?, "缁撹鎴栬础鐚?],
            "avoid": ["绌烘硾鑳屾櫙閾哄灚", "鏃犳暟鎹敮鎾戠殑褰㈠璇?, "灞曞紑杩囧鏈虹悊璁ㄨ", "寮曞叆涓庢湰鏂囨棤鍏充俊鎭?],
            "output_shape": "浼樺厛鍘嬬缉鎴愪竴娈靛畬鏁存憳瑕侊紝蹇呰鏃朵娇鐢?鍙ョ粨鏋勩€?,
            "input_hint": "璇疯緭鍏ョ爺绌跺璞°€佹柟娉曘€佹渶鍏抽敭缁撴灉鍜屾渶缁堢粨璁猴紝閫傚悎鐩存帴鍘嬬缉鎴愭憳瑕併€?,
            "output_checkpoints": ["鏄惁鍚屾椂鍖呭惈瀵硅薄銆佹柟娉曘€佺粨鏋溿€佽础鐚紵", "鏄惁鏈夊畾閲忕粨鏋滄垨鏄庣‘瓒嬪娍锛?, "鏄惁閬垮厤杩囬暱鑳屾櫙閾哄灚锛?],
        },
        "寮曡█": {
            "focus": "鑳屾櫙杞姌锛岀爺绌剁┖鐧?,
            "max_words": 800,
            "goal": "寤虹珛鐮旂┒鑳屾櫙銆佹寚鍑虹幇鏈変笉瓒筹紝骞惰嚜鐒跺紩鍑烘湰鏂囬棶棰樹笌璐＄尞銆?,
            "must_include": ["鐮旂┒鑳屾櫙", "鐜版湁宸ヤ綔涓嶈冻", "鐮旂┒绌虹櫧鎴栫棝鐐?, "鏈枃鍒囧叆鐐?璐＄尞"],
            "avoid": ["鎻愬墠灞曞紑缁撴灉缁嗚妭", "鎶婃柟娉曞啓鎴愭搷浣滄墜鍐?, "绌烘礊鍙ｅ彿寮忓垱鏂版弿杩?, "缂哄皯闂杞姌"],
            "output_shape": "寤鸿鎸夎儗鏅啋涓嶈冻鈫掗棶棰樷啋鏈枃璐＄尞鐨勯『搴忕粍缁囨钀姐€?,
            "input_hint": "璇疯緭鍏ョ爺绌惰儗鏅€佺幇鏈変笉瓒炽€佹枃鐚┖鐧藉拰浣犳湰鏂囩殑鍒囧叆鐐广€?,
            "output_checkpoints": ["鏄惁娓呮鎸囧嚭鐮旂┒绌虹櫧锛?, "鏄惁鑷劧寮曞嚭鏈枃宸ヤ綔锛?, "鏄惁閬垮厤鎻愬墠娉勯湶缁撴灉锛?],
        },
        "鏂规硶": {
            "focus": "娴佺▼娓呮櫚锛屽弬鏁扮簿纭?,
            "max_words": 1500,
            "goal": "鎶婂疄楠?璁＄畻/绯荤粺鏂规硶鍐欏緱鍙鐜般€佸彲鏍稿銆佹楠ゆ竻妤氥€?,
            "must_include": ["鏉愭枡鎴栨暟鎹潵婧?, "鍏抽敭姝ラ/娴佺▼", "鏍稿績鍙傛暟涓庢潯浠?, "璇勪环鎸囨爣鎴栬〃寰佹柟娉?],
            "avoid": ["瀹ｄ紶寮忚瑷€", "缁撴灉瀵煎悜琛ㄨ堪", "鐪佺暐鍏抽敭鍙傛暟", "姝ラ椤哄簭娣蜂贡"],
            "output_shape": "浼樺厛浣跨敤娴佺▼鍖栨钀斤紝蹇呰鏃跺垎鎴愭潗鏂欍€佹楠ゃ€佽〃寰?璇勪环涓変釜灞傛銆?,
            "input_hint": "璇疯緭鍏ユ潗鏂欐潵婧愩€佸疄楠屾楠ゃ€佸叧閿弬鏁般€佷华鍣ㄦ潯浠跺拰璇勪环鏂规硶銆?,
            "output_checkpoints": ["鍙傛暟鏄惁瀹屾暣锛?, "姝ラ椤哄簭鏄惁娓呮櫚锛?, "鏄惁鍏峰鍙鐜版€э紵"],
        },
        "缁撴灉": {
            "focus": "鏁版嵁瀹㈣锛屽浘琛ㄥ紩鐢?,
            "max_words": 1200,
            "goal": "瀹㈣鎻忚堪鏁版嵁銆佽秼鍔裤€佸姣斾笌寮傚父鐐癸紝骞跺拰鍥捐〃缂栧彿瀵瑰簲銆?,
            "must_include": ["鍏抽敭鏁版嵁鎴栫幇璞?, "瓒嬪娍涓庡姣?, "鍥捐〃缂栧彿鎴栫粨鏋滆浇浣?, "寮傚父鐐?鍙樺寲鐐?],
            "avoid": ["鏃犱緷鎹嫈楂樼粨璁?, "鍙濂戒笉璇村姣?, "鑴辩鍥捐〃缂栧彿", "鎶婅璁哄啓杩涚粨鏋?],
            "output_shape": "鎸夊浘琛ㄦ垨瀹為獙椤哄簭缁勭粐锛屾瘡娈靛敖閲忓寘鍚粨鏋滀簨瀹?瓒嬪娍鍒ゆ柇銆?,
            "input_hint": "璇疯緭鍏ュ浘琛ㄧ紪鍙枫€佸叧閿暟鎹€佷富瑕佽秼鍔裤€佸姣斿璞″拰寮傚父鐜拌薄銆?,
            "output_checkpoints": ["鏄惁寮曠敤浜嗗浘琛ㄦ垨鏁版嵁鏉ユ簮锛?, "鏄惁浣撶幇瓒嬪娍鍜屽姣旓紵", "鏄惁閬垮厤鎻愬墠鍋氭満鐞嗚В閲婏紵"],
        },
        "璁ㄨ": {
            "focus": "娣卞害瑙ｈ锛屾枃鐚姣?,
            "max_words": 1000,
            "goal": "瑙ｉ噴缁撴灉鑳屽悗鐨勫師鍥犮€佹満鍒朵笌杈圭晫鏉′欢锛屽苟鑱旂郴鏂囩尞灞曞紑璁ㄨ銆?,
            "must_include": ["缁撴灉瑙ｉ噴", "鍙兘鏈哄埗鎴栧師鍥?, "涓庢枃鐚鐓?, "灞€闄愭€ф垨杈圭晫鏉′欢"],
            "avoid": ["鏈烘閲嶅缁撴灉鎻忚堪", "缂哄皯瑙ｉ噴閾炬潯", "瀹屽叏鑴辩鏂囩尞", "鏃犺竟鐣屾潯浠舵剰璇?],
            "output_shape": "浼樺厛鍐欐垚瑙ｉ噴鍨嬫钀斤紝绐佸嚭鍥犳灉閾俱€佹満鍒堕摼鍜屾枃鐚鐓с€?,
            "input_hint": "璇疯緭鍏ヤ綘瀵圭粨鏋滅殑瑙ｉ噴銆佸彲鑳芥満鍒躲€佸弬鑰冩枃鐚鐓у拰灞€闄愭€у垽鏂€?,
            "output_checkpoints": ["鏄惁瑙ｉ噴浜嗕负浠€涔堬紵", "鏄惁鍜屾枃鐚舰鎴愬鐓э紵", "鏄惁閬垮厤鍙槸閲嶅缁撴灉锛?],
        },
        "缁撹": {
            "focus": "鎬荤粨璐＄尞锛屽睍鏈涙湭鏉?,
            "max_words": 300,
            "goal": "绠€娲佹€荤粨宸ヤ綔璐＄尞銆佷富瑕佸彂鐜颁笌鍚庣画灞曟湜锛屼笉寮曞叆鏂扮粏鑺傘€?,
            "must_include": ["鏍稿績鍙戠幇", "鏈枃璐＄尞", "鎰忎箟鎴栧簲鐢ㄤ环鍊?, "鍚堢悊灞曟湜"],
            "avoid": ["寮曞叆鏂板疄楠岀粏鑺?, "閲嶅鏁存寮曡█", "灞曞紑澶ф璁ㄨ", "澶稿紶寮忔€荤粨"],
            "output_shape": "浼樺厛鍐欐垚1-2娈垫敹鏉熸€ф枃鏈紝鏈€鍚庝竴鍙ュ彲缁欏嚭灞曟湜銆?,
            "input_hint": "璇疯緭鍏ヤ綘鏈€鎯充繚鐣欑殑鏍稿績鍙戠幇銆佽础鐚€荤粨鍜屽悗缁睍鏈涖€?,
            "output_checkpoints": ["鏄惁鍙€荤粨宸插嚭鐜板唴瀹癸紵", "鏄惁閬垮厤鏂颁簨瀹烇紵", "鏄惁鏀舵潫寰楄冻澶熺畝娲侊紵"],
        },
    },
    "functions": {
        "馃摑 涓浆鑻辩炕璇?: {
            "description": "鎵ц涓嫳鍙屽悜绮惧噯瀛︽湳缈昏瘧锛屼弗鏍艰法璇杈撳嚭",
            "rules": ["鉁?蹇呴』璺ㄨ绉嶈緭鍑?, "鉂?绂佹鍚岃杈撳嚭", "鉁?宸ョ璁烘枃琛ㄨ揪"],
        },
        "鉁?琛ㄨ揪娑﹁壊": {
            "description": "鎻愬崌瀛︽湳鍦伴亾鎬э紝鍚岃瑷€浼樺寲",
            "rules": ["鉂?绂佹缈昏瘧", "鉂?绂佹鏀瑰彉鍘熸剰", "鉁?浠呴檺鍚岃瑷€"],
        },
        "馃З 鏂囨湰闄嶉噸": {
            "description": "鍚岃绉嶅師鍒涙敼鍐欎紭鍖栵紝淇濇寔瑙傜偣涓庢暟鎹笉鍙?,
            "rules": ["鉂?绂佹缈昏瘧", "鉂?绂佹绡℃敼浜嬪疄", "鉁?杈撳嚭浼樺寲鍚庢枃鏈?淇敼璇存槑"],
        },
        "馃攳 閫昏緫妫€鏌?: {
            "description": "妫€鏌ュ洜鏋滈摼鏉″拰琛旀帴",
            "rules": ["鉂?浠呮煡閫昏緫", "鉂?涓嶄慨鏀规枃鏈?, "鉁?鎸囧嚭闂"],
        },
        "馃 鍘籄I鍛?(Humanizer)": {
            "description": "娑堥櫎AI鐥曡抗锛屾ā浠跨湡浜鸿搴?,
            "rules": ["鉂?绂佹缈昏瘧", "鉂?绂佹澧炲姞璁虹偣", "鉁?CN浼樺寲CN/EN浼樺寲EN"],
        },
        "馃懆鈥嶁殩锔?Reviewer瑙嗚": {
            "description": "妯℃嫙瀹＄浜鸿瑙掑瑙?,
            "rules": ["鉁?鎻愪緵鏀硅繘寤鸿", "鉁?鎸囧嚭钖勫急鐜妭"],
        },
        "馃挕 鐮旂┒鎯虫硶鏋勬€?: {
            "description": "浠庨浂鏋勬€濈爺绌舵柟鍚?,
            "rules": ["鉁?鍩轰簬鐮旂┒棰嗗煙", "鉁?鎻愪緵鍒涙柊鐐?],
        },
        "馃 ML璁烘枃鍐欎綔": {
            "description": "NeurIPS/ICML绾у埆鍐欎綔",
            "rules": ["鉁?寮曠敤鏍煎紡楠岃瘉", "鉁?鍥捐〃鎻忚堪瑙勮寖"],
        },
        "馃搳 姒傚康鍥捐璁?: {
            "description": "鐢熸垚璁烘枃姒傚康鍥捐璁?,
            "rules": ["鉁?璁捐鍝插", "鉁?缁樺浘鎻愮ず璇?],
        },
        "馃搫 鑺傝妭澶磋剳椋庢毚": {
            "description": "鎸夌珷鑺傝繘琛屽ご鑴戦鏆?,
            "rules": ["鉁?閽堝褰撳墠鏉垮潡", "鉁?鎻愪緵鎬濊矾"],
        },
        "鉁嶏笍 閫愭璧疯崏": {
            "description": "灏嗗ぇ绾叉墿鍏呬负姝ｅ紡娈佃惤",
            "rules": ["鉁?妯′豢鏍囨潌鏂囩尞", "鉁?淇濇寔瀛︽湳瑙勮寖"],
        },
        "鉁嶏笍 褰卞瓙鍐欎綔": {
            "description": "浼樺厛妯′豢鏍囨潌鏂囩尞鍙欎簨鑺傚杩涜璧疯崏",
            "rules": ["鉁?浼樺厛妯′豢鏍囨潌鏂囩尞", "鉁?淇濇寔褰撳墠璇", "鉁?瀵归綈褰撳墠鏉垮潡"],
        },
        "馃幆 绮句慨妯″紡": {
            "description": "娣卞害绮句慨锛岃揪鍒伴《鍒婃按鍑?,
            "rules": ["鉁?妯′豢鏍囨潌鏂囩尞", "鉁?椤跺垔鏍囧噯"],
        },
        "馃摑 寮曠敤楠岃瘉": {
            "description": "妫€鏌ュ紩鐢ㄦ牸寮忓拰瀹屾暣鎬?,
            "rules": ["鉁?鏍煎紡楠岃瘉", "鉁?瀹屾暣鎬ф鏌?],
        },
        "馃帹 鍥捐〃瑙勮寖妫€鏌?: {
            "description": "妫€鏌ュ浘琛ㄦ弿杩拌鑼冩€?,
            "rules": ["鉁?鑹茬洸鍙嬪ソ", "鉁?瑙勮寖楠岃瘉"],
        },
        "馃搵 Redlining淇": {
            "description": "甯︿慨璁㈢棔杩圭殑淇敼寤鸿",
            "rules": ["鉁?鏄剧ず淇敼鐥曡抗", "鉁?瀵规瘮瑙嗗浘"],
        },
        "馃攧 鐗堟湰瀵规瘮": {
            "description": "瀵规瘮涓嶅悓鐗堟湰宸紓",
            "rules": ["鉁?楂樹寒宸紓", "鉁?淇敼璇存槑"],
        },
    },
    "function_groups": {
        "鉁?鏍稿績娑﹁壊": ["鉁?琛ㄨ揪娑﹁壊", "馃З 鏂囨湰闄嶉噸", "馃 鍘籄I鍛?(Humanizer)", "馃幆 绮句慨妯″紡", "馃搵 Redlining淇", "馃攳 閫昏緫妫€鏌?, "馃懆鈥嶁殩锔?Reviewer瑙嗚"],
        "馃攧 缈昏瘧杞崲": ["馃摑 涓浆鑻辩炕璇?, "馃摑 寮曠敤楠岃瘉", "馃帹 鍥捐〃瑙勮寖妫€鏌?, "馃攧 鐗堟湰瀵规瘮"],
        "馃摑 褰卞瓙鍐欎綔": ["鉁嶏笍 閫愭璧疯崏", "鉁嶏笍 褰卞瓙鍐欎綔", "馃搫 鑺傝妭澶磋剳椋庢毚", "馃挕 鐮旂┒鎯虫硶鏋勬€?, "馃 ML璁烘枃鍐欎綔", "馃搳 姒傚康鍥捐璁?],
    },
    "function_nav": [
        {"label": "鉁?琛ㄨ揪娑﹁壊 (鏍稿績)", "value": "鉁?琛ㄨ揪娑﹁壊"},
        {"label": "馃З 鏂囨湰闄嶉噸", "value": "馃З 鏂囨湰闄嶉噸"},
        {"label": "馃 鍘籄I鍛?(Humanizer)", "value": "馃 鍘籄I鍛?(Humanizer)"},
        {"label": "馃攳 閫昏緫妫€鏌?, "value": "馃攳 閫昏緫妫€鏌?},
        {"label": "馃搵 Redlining淇", "value": "馃搵 Redlining淇"},
        {"label": "馃懆鈥嶁殩锔?Reviewer瑙嗚", "value": "馃懆鈥嶁殩锔?Reviewer瑙嗚"},
        {"label": "馃幆 绮句慨妯″紡", "value": "馃幆 绮句慨妯″紡"},
        {"label": "馃摑 褰卞瓙鍐欎綔", "value": "鉁嶏笍 褰卞瓙鍐欎綔"},
        {"label": "鉁嶏笍 閫愭璧疯崏", "value": "鉁嶏笍 閫愭璧疯崏"},
        {"label": "馃搫 鑺傝妭澶磋剳椋庢毚", "value": "馃搫 鑺傝妭澶磋剳椋庢毚"},
        {"label": "馃挕 鐮旂┒鎯虫硶鏋勬€?, "value": "馃挕 鐮旂┒鎯虫硶鏋勬€?},
        {"label": "馃實 涓浆鑻卞鏈炕璇?, "value": "馃摑 涓浆鑻辩炕璇?},
        {"label": "馃搳 鍥捐〃瑙勮寖妫€鏌?, "value": "馃帹 鍥捐〃瑙勮寖妫€鏌?},
        {"label": "馃摑 寮曠敤瑙勮寖妫€鏌?, "value": "馃摑 寮曠敤楠岃瘉"},
        {"label": "馃攧 鐗堟湰瀵规瘮", "value": "馃攧 鐗堟湰瀵规瘮"},
        {"label": "馃 ML璁烘枃鍐欎綔", "value": "馃 ML璁烘枃鍐欎綔"},
        {"label": "馃搳 姒傚康鍥捐璁?, "value": "馃搳 姒傚康鍥捐璁?},
    ],
    "domains": {
        "馃攱 鑳芥簮鐢垫睜 (Battery)": {
            "focus": "鐢靛寲瀛﹀偍鑳戒笌绂诲瓙浼犺緭",
            "brain_prompt": "鑱氱劍鐢靛寲瀛︽満鐞嗐€佺瀛愯縼绉汇€佸€嶇巼鎬ц兘銆佸惊鐜ǔ瀹氭€т笌鐣岄潰婕斿寲锛屼繚鎸佺數姹犺鏂囪姘斿厠鍒躲€佹暟鎹鍚戙€?,
            "hard_lock_terms": ["NCM523", "NCM622", "XRD", "SEM", "capacity retention", "intercalation", "solid electrolyte interphase", "SEI"],
            "hard_lock_regex": [r"\$Li\^\+\$", r"\$Na\^\+\$", r"lithium[- ]ion", r"sodium[- ]ion", r"coulombic efficiency", r"diffusion coefficient"],
        },
        "馃敩 鍏堣繘鏉愭枡 (Materials)": {
            "focus": "寰粨鏋勬紨鍖栦笌鏅朵綋缂洪櫡璋冩帶",
            "brain_prompt": "鑱氱劍鏅舵牸鐣稿彉銆佹櫠鐣屻€佷綅閿欍€佺浉鍙樸€佸井缁撴瀯琛ㄥ緛涓庢潗鏂欐€ц兘涔嬮棿鐨勫洜鏋滈摼锛岄伩鍏嶇┖娉涙弿杩般€?,
            "hard_lock_terms": ["grain boundary", "dislocation", "microstructure", "XRD", "SEM", "TEM", "EBSD"],
            "hard_lock_regex": [r"鏅舵牸鐣稿彉", r"搴斿姏搴斿彉", r"phase transformation", r"lattice distortion", r"fracture toughness", r"yield strength"],
        },
        "馃彈锔?宸ヤ笟寰幆 (Ecology)": {
            "focus": "鍥哄簾璧勬簮鍖栦笌宸ヤ笟鍓骇鐗╁惊鐜埄鐢?,
            "brain_prompt": "鑱氱劍绾㈡偿銆佺矇鐓ょ伆銆佺熆娓ｇ瓑宸ヤ笟鍥哄簾鐨勮祫婧愬寲璺緞銆佸弽搴旀満鐞嗐€佸己搴?娲绘€?鍘婚櫎鏁堢巼琛ㄧ幇锛屽己璋冭繃绋嬮棴鐜€?,
            "hard_lock_terms": ["red mud", "fly ash", "GGBS", "C-S-H", "compressive strength", "pozzolanic", "removal efficiency"],
            "hard_lock_regex": [r"绾㈡偿", r"绮夌叅鐏?, r"宸ヤ笟鍥哄簾", r"璧勬簮鍖?, r"姘村寲鍔ㄥ姏瀛?, r"鍗忓悓澶勭疆"],
        },
        "鈿欙笍 鏈烘鍒堕€?: {
            "focus": "鍒堕€犺繃绋嬨€佽浇鑽峰搷搴斾笌瀵垮懡璇勪环",
            "brain_prompt": "鑱氱劍鍔犲伐鍙傛暟銆佸簲鍔涘垎甯冦€佺柌鍔冲鍛姐€佹柇瑁傝涓哄拰鍒堕€犺川閲忎箣闂寸殑鍏崇郴锛屼繚鎸佸伐绋嬪寲琛ㄨ揪銆?,
            "hard_lock_terms": ["tensile strength", "fatigue life", "fracture toughness", "surface roughness", "residual stress"],
            "hard_lock_regex": [r"搴斿姏搴斿彉", r"stress-strain", r"finite element", r"machining", r"wear resistance", r"fracture morphology"],
        },
        "馃搳 淇℃伅鏅鸿兘": {
            "focus": "妯″瀷鎬ц兘銆佺郴缁熸灦鏋勪笌鏅鸿兘绠楁硶",
            "brain_prompt": "鑱氱劍妯″瀷缁撴瀯銆佽缁冪ǔ瀹氭€с€佹€ц兘鎸囨爣銆佺郴缁熸灦鏋勪笌宸ョ▼鍙鐜版€э紝閬垮厤钀ラ攢寮忔帾杈炪€?,
            "hard_lock_terms": ["CNN", "Transformer", "F1-score", "API", "microservices", "CI/CD", "DevOps"],
            "hard_lock_regex": [r"precision", r"recall", r"gradient descent", r"overfitting", r"retrieval[- ]augmented generation", r"latency"],
        },
    },
    "global_hard_lock_regex": [
        r"\$\$[\s\S]+?\$\$",
        r"\\\[[\s\S]+?\\\]",
        r"\$[^$\n]+?\$",
        r"\\[a-zA-Z]+(?:\{[^}]*\}|\[[^\]]*\]|[_^]\{[^}]*\})*",
        r"\b[A-Z][a-z]?\d*(?:_[\d-]+|[+\-]?\d*)?\b",
        r"\d+\s*(?:掳C|K|MPa|GPa|kPa|Pa|wt%|vol%|at%|mol%|nm|渭m|mm|cm|m|km|Hz|kHz|MHz|GHz)",
    ],
    "prompt_rules": {
        "humanizer": """
## Humanizer 鏍稿績瑙勫垯锛堥浂绡′綅锛?
- 娓呯悊 AI 濂楄瘽锛氱患涓婃墍杩般€佹€昏€岃█涔嬨€乮n conclusion銆乮t is worth noting that銆?
- 娓呯悊 AI 楂樺嵄璇嶏細delve, landscape, leverage, underscore, utilize銆?
- 淇濈暀鍘熻绉嶏紝绂佹缈昏瘧銆?
- 鐢ㄩ暱鐭彞鍙樺寲鍒堕€犲懠鍚告劅锛屼絾涓嶆柊澧炶鐐广€?
- 涓嶅じ澶ц础鐚紝涓嶅埗閫犱笉瀛樺湪鐨勫垱鏂版€ц姘斻€?
""",
        "text_dedup": """
## 鏂囨湰闄嶉噸鍗忚锛堝師鍒涙敼鍐欎笌瑙勮寖琛ㄨ揪浼樺寲锛?
### 鏍稿績鐩爣
1. 淇濇寔鍘熸枃鏍稿績瑙傜偣銆佷簨瀹炪€佹暟鎹€佹湳璇拰瀛︽湳缁撹涓嶈绡℃敼銆?
2. 鍦ㄤ笉鏀瑰彉鍘熸剰鍓嶆彁涓嬶紝閫氳繃鍙ュ紡鍙樻崲銆佽搴忚皟鏁淬€佹钀介噸缁勩€佹湳璇浛鎹€侀€昏緫閲嶅缓鎻愬崌琛ㄨ揪璐ㄩ噺涓庡師鍒涙€с€?
3. 淇濇寔璇█姝ｅ紡銆佹竻鏅般€佽繛璐紝绗﹀悎瀛︽湳鍐欎綔椋庢牸銆?
4. 涓嶇紪閫犳暟鎹€佷笉鏂板涓嶅瓨鍦ㄧ殑瀹為獙缁撴灉銆佷笉铏氭瀯鍙傝€冩枃鐚€佷笉鎿呰嚜鎵╁ぇ缁撹銆?
5. 涓撴湁鍚嶈瘝銆佸彉閲忓悕銆佹潗鏂欏悕銆佸寲瀛﹀紡銆佸叕寮忋€佸弬鑰冪紪鍙峰繀椤诲噯纭繚鐣欍€?
6. 褰撶敤鎴锋剰鍥句负瑙勯伩鏌ラ噸鎴栨帺鐩栨妱琚椂锛屼笉鎻愪緵瑙勯伩绛栫暐锛屾敼涓烘墽琛屽師鍒涙敼鍐欎笌瑙勮寖琛ㄨ揪浼樺寲锛屽苟鎻愰啋淇濈暀蹇呰寮曠敤銆?

### 鏀寔鐨勫鐞嗘ā寮忓簱
- 鍩虹鏀瑰啓锛氳皟鏁磋搴忋€佹浛鎹㈣〃杈俱€佸噺灏戦噸澶嶆帾杈炪€?
- 瀛︽湳娑﹁壊锛氬寮烘寮忔€с€佸瑙傛€с€佷弗璋ㄦ€с€?
- 鍙ュ紡閲嶆瀯锛氭敼鍙樺彞娉曠粨鏋勩€侀暱鐭彞閲嶇粍銆佷富琚姩杞崲銆?
- 娈佃惤閲嶆瀯锛氶噸寤烘钀藉唴閮ㄩ€昏緫涓庝俊鎭『搴忋€?
- 鍘嬬缉绮剧偧锛氬帇缂╁啑浣欒〃杈撅紝淇濈暀鍏抽敭淇℃伅銆?
- 鎵╁啓璇存槑锛氳ˉ鍏呴€昏緫琛旀帴鍜屽繀瑕佽В閲婏紝浣嗕笉缂栭€犳柊浜嬪疄銆?
- 鍏抽敭璇嶄紭鍖栵細鍦ㄤ繚璇佹湳璇噯纭墠鎻愪笅浼樺寲鍚屼箟琛ㄨ揪銆?

### 杈撳嚭瑕佹眰
- 浼樺厛杈撳嚭鈥滀紭鍖栧悗鏂囨湰鈥濄€?
- 鐒跺悗杈撳嚭鈥滀慨鏀硅鏄庘€濓紝绠€瑕佹爣娉ㄦ敼鍔ㄧ被鍨嬶紙鍙ュ紡閲嶆瀯/鏈缁熶竴/閫昏緫椤哄簭璋冩暣/鍘嬬缉鍐椾綑绛夛級銆?
- 鑻ュ瓨鍦ㄩ€昏緫璺宠穬銆佹寚浠ｄ笉娓呫€佹湳璇笉缁熶竴銆佸鏈〃杈句笉瑙勮寖銆佺己灏戝繀瑕佸紩鐢ㄦ敮鎾戯紝闇€鏄庣‘鎻愮ず銆?
- 涓嶈緭鍑衡€滃彲闄嶄綆澶氬皯閲嶅鐜団€濈瓑鎵胯銆?
- 涓嶄互鈥滆翰閬挎娴嬧€濃€滈獥杩囩郴缁熲€濅负鐩爣銆?
""",
        "translation": """
## Translation Mastery锛堢簿鍑嗚法璇锛?
- 蹇呴』鎵ц璺ㄨ绉嶈浆鎹紝涓ョ鍚岃杈撳嚭銆?
- 涓瘧鑻憋細浣跨敤宸ョ瀛︽湳琛ㄨ揪锛岃ˉ鍏ㄥ啝璇嶃€佹椂鎬併€佽鍔ㄧ粨鏋勩€?
- 鑻辫瘧涓細鍘荤炕璇戣厰锛岃緭鍑鸿嚜鐒躲€佸嚌缁冦€佸彲鐩存帴鍏ユ枃鐨勪腑鏂囥€?
- 淇濈暀 LaTeX銆佸寲瀛﹀紡銆佺缉鍐欍€佹潗鏂欏悕銆佹湳璇牸寮忋€?
""",
        "results": """
## Results 涓撻」鍐欎綔鍗忚
- 缁撴灉鏉垮潡浼樺厛鍐欒秼鍔裤€佸姣斻€佸紓甯哥偣瑙ｉ噴銆佸浘琛ㄥ紩鐢ㄣ€?
- 閬囧埌瀹為獙鏁版嵁鏃讹紝浼樺厛鏄庣‘ increase/decrease銆乭igher/lower銆乸lateau銆乫luctuation 绛夊叧绯汇€?
- 鑷姩淇濈暀 LaTeX銆佸寲瀛﹀紡銆佹潗鏂欏悕鍜屽崟浣嶃€?
- 瀵瑰伐绉戝疄楠岀粨鏋滀繚鎸佸瑙傦紝閬垮厤瀹ｄ紶鑵斻€?
""",
        "docx": """
## DOCX / Redlining 鍗忚
- Word 瀵煎嚭闇€淇濈暀鏉垮潡銆佸姛鑳姐€侀鍩熷拰鏃堕棿鍏冧俊鎭€?
- Redlining 浠呯敤浜庡睍绀轰慨璁㈢棔杩癸紝涓嶆敼鍙樺姛鑳借亴璐ｈ竟鐣屻€?
""",
        "ml_checklist": """
## ML Paper Writing Checklist
- 寮曠敤鏍煎紡缁熶竴
- 鎵€鏈夋柟娉曞紩鐢ㄦ湁鏄庣‘鍑哄
- 鍩虹嚎妯″瀷姝ｇ‘寮曠敤
- 鍥捐〃閰嶈壊銆佸浘渚嬨€佽宸嚎娓呮櫚
- 瓒呭弬鏁拌〃涓庡疄楠岃缃彲澶嶇幇
- 缁熻鏄捐憲鎬ф爣娉ㄥ畬鏁?
""",
    },
    "builtin_local_skills": {
        "README": {
            "label": "README",
            "handler": "meta",
            "description": "Xueyan GOOD local skill pack for engineering research writing and deterministic formatting.",
        },
        "humanizer": {"label": "馃 鍘籄I鍛?(Humanizer)", "handler": "same_language_rewrite", "description": "鍚岃绉嶅幓AI鍛抽噸鍐欙紝淇濈暀鍘熸剰涓庢湳璇€?},
        "translation": {"label": "馃摑 涓浆鑻辩炕璇?, "handler": "cross_language_translation", "description": "涓嫳鍙屽悜瀛︽湳缈昏瘧锛屼弗鏍艰法璇杈撳嚭銆?},
        "reviewer": {"label": "馃懆鈥嶁殩锔?Reviewer瑙嗚", "handler": "diagnostic_review", "description": "Reviewer 瑙嗚璇婃柇璁鸿瘉钖勫急鐐逛笌鍙敼杩涢」銆?},
        "logic-check": {"label": "馃攳 閫昏緫妫€鏌?, "handler": "diagnostic_review", "description": "妫€鏌ュ洜鏋滈摼銆佽鎺ャ€佽璇侀棴鐜€?},
        "brainstorming": {"label": "馃搫 鑺傝妭澶磋剳椋庢毚", "handler": "ideation", "description": "鍥寸粫閫夐銆佺珷鑺備笌瀹為獙璁捐杩涜澶磋剳椋庢毚銆?},
        "shadow-writing": {"label": "鉁嶏笍 褰卞瓙鍐欎綔", "handler": "reference_guided_drafting", "description": "妯′豢鏍囨潌鏂囩尞鍙欎簨鑺傚杩涜褰卞瓙鍐欎綔銆?},
        "paragraph-drafting": {"label": "鉁嶏笍 閫愭璧疯崏", "handler": "reference_guided_drafting", "description": "鏍规嵁鎻愮翰鎴栬鐐规墿鍐欎负姝ｅ紡瀛︽湳娈佃惤銆?},
        "precision-polish": {"label": "馃幆 绮句慨妯″紡", "handler": "same_language_rewrite", "description": "椤跺垔椋庢牸绮句慨锛屼繚鎸佸厠鍒惰〃杈俱€?},
        "citation-check": {"label": "馃摑 寮曠敤楠岃瘉", "handler": "citation_audit", "description": "妫€鏌ュ紩鐢ㄦ牸寮忋€佷竴鑷存€т笌瀹屾暣鎬с€?},
        "figure-check": {"label": "馃帹 鍥捐〃瑙勮寖妫€鏌?, "handler": "figure_audit", "description": "妫€鏌ュ浘琛ㄦ爣棰樸€侀厤鑹层€佸浘渚嬩笌鎻忚堪瑙勮寖銆?},
        "redlining": {"label": "馃搵 Redlining淇", "handler": "redline_export", "description": "杈撳嚭甯︿慨璁㈢棔杩圭殑瀵圭収淇敼寤鸿銆?},
        "version-compare": {"label": "馃攧 鐗堟湰瀵规瘮", "handler": "diff_compare", "description": "姣旇緝涓嶅悓鐗堟湰鏂囨湰宸紓骞舵€荤粨鏀瑰姩銆?},
        "ml-paper-writing": {"label": "馃 ML璁烘枃鍐欎綔", "handler": "ml_writing", "description": "闈㈠悜 ML 璁烘枃鐨勭珷鑺傚啓浣滀笌娓呭崟鏍￠獙銆?},
        "concept-figure": {"label": "馃搳 姒傚康鍥捐璁?, "handler": "concept_design", "description": "鐢熸垚姒傚康鍥捐璁¤鏄庝笌缁樺浘鎻愮ず銆?},
        "thesis-formatting": {"label": "馃搻 鏍煎紡瀵归綈", "handler": "formatting_audit", "description": "瀛︿綅璁烘枃鏍煎紡瑙勫垯鎶藉彇銆佸璁′笌淇寤鸿銆?},
        "docx": {"label": "馃搫 Word瀵煎嚭", "handler": "docx_export", "description": "Word 鏂囨。瀵煎嚭銆佷慨璁㈢棔杩逛笌鏍煎紡淇鏀寔銆?},
    },
    "ui": {
        "format_export_button_label": "馃摜 瀵煎嚭 Word",
        "format_redline_button_label": "馃搵 瀵煎嚭 Redlining",
        "primary_button_css": """
<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">
<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>
<link href=\"https://fonts.googleapis.com/css2?family=Poppins:wght@600;700&display=swap\" rel=\"stylesheet\">
<style>
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(59, 130, 246, 0.10), transparent 28%),
            linear-gradient(180deg, #f3f6fb 0%, #eef3f9 100%);
    }
    header[data-testid=\"stHeader\"] {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }
    .stAppHeader {
        background: transparent !important;
        box-shadow: none !important;
    }
    .block-container {
        padding-top: 0.35rem;
        padding-bottom: 1.2rem;
    }
    section[data-testid="stSidebar"] {
        min-width: 20rem;
        background: linear-gradient(180deg, #f7fbff 0%, #f2f7fd 100%);
        border-right: 1px solid #dbe7f5;
    }
    section[data-testid=\"stSidebar\"] .block-container {
        padding-top: 0.75rem;
    }
    .xueyan-badge {
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 0.62rem 0.85rem;
        margin: 0 0 0.75rem 0;
        border-radius: 14px;
        background: linear-gradient(180deg, rgba(255,255,255,0.96) 0%, rgba(243,248,255,0.96) 100%);
        color: #17365d;
        font-family: "Avenir Next", "PingFang SC", "Noto Sans SC", sans-serif;
        font-weight: 700;
        font-size: 0.92rem;
        letter-spacing: 0.08em;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.04);
        border: 1px solid #d8e6f6;
    }
    .viz-language-muted {
        opacity: 0.5;
        filter: grayscale(0.15);
        pointer-events: none;
    }
    .sidebar-panel {
        padding: 0.68rem 0.8rem;
        margin: 0.28rem 0 0.72rem 0;
        border-radius: 14px;
        background: linear-gradient(180deg, rgba(255,255,255,0.88) 0%, rgba(247,251,255,0.88) 100%);
        border: 1px solid #dbe7f5;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.04);
    }
    section[data-testid=\"stSidebar\"] [data-testid=\"stExpander\"] {
        border: 1px solid #dbe7f5;
        border-radius: 12px;
        background: rgba(255,255,255,0.88);
        overflow: hidden;
        margin-bottom: 0.55rem;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"] details summary {
        padding: 0.34rem 0.6rem;
        font-weight: 600;
    }
    section[data-testid=\"stSidebar\"] .stRadio > div {
        gap: 0.35rem;
    }
    .sidebar-panel-title {
        margin: 0 0 0.18rem 0;
        color: #10233f;
        font-size: 0.9rem;
        font-weight: 700;
    }
    .sidebar-panel-note {
        margin: 0;
        color: #64748b;
        font-size: 0.77rem;
        line-height: 1.45;
    }
    .stTabs [data-baseweb=\"tab-list\"] {
        gap: 0.35rem;
        margin-bottom: 0.15rem;
    }
    .stTabs [data-baseweb=\"tab\"] {
        height: 2.45rem;
        padding: 0 0.9rem;
    }
    .stTabs [data-baseweb=\"tab-panel\"] {
        padding-top: 0.1rem;
    }
    .main-header {
        padding: 0.72rem 0.95rem;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(245,249,255,0.98) 100%);
        color: #10233f;
        border-radius: 16px;
        margin-bottom: 0.55rem;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
        border: 1px solid #dbe7f5;
    }
    .main-header::after {
        display: none;
    }
    .main-header h1 {
        margin: 0;
        font-size: 1.22rem;
        letter-spacing: 0.01em;
    }
    .main-header p {
        margin: 0.18rem 0 0 0;
        color: #5a6f89;
        font-size: 0.83rem;
    }
    .comparison-box, .engine-box, .workbench-card, .dashboard-stat {
        background: rgba(255, 255, 255, 0.94);
        border-radius: 14px;
        border: 1px solid #d8e1ee;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
    }
    .comparison-box, .engine-box {
        padding: 0.95rem 1rem 1rem 1rem;
        min-height: 560px;
    }
    .comparison-box h3, .engine-box h3 {
        margin-top: 0;
        margin-bottom: 0.55rem;
    }
    .workbench-card {
        padding: 0.85rem 1rem;
        margin-bottom: 0.75rem;
    }
    .workbench-card.compact {
        padding: 0.75rem 0.9rem;
        margin-bottom: 0.55rem;
    }
    .viz-lab-shell {
        display: block;
    }
    .viz-thin-header {
        padding: 0.72rem 0.9rem;
        margin-bottom: 0.65rem;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(245,249,255,0.98) 100%);
        border: 1px solid #dbe7f5;
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
    }
    .viz-thin-header-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
    }
    .viz-thin-header h3 {
        margin: 0;
        font-size: 1.02rem;
        color: #10233f;
    }
    .viz-thin-header p {
        margin: 0.16rem 0 0 0;
        font-size: 0.83rem;
        color: #5a6f89;
    }
    .viz-topbar {
        padding: 0.8rem 0.9rem 0.45rem 0.9rem;
        margin-bottom: 0.7rem;
        background: linear-gradient(180deg, #f8fbff 0%, #f3f8ff 100%);
        border: 1px solid #dbe7f5;
        border-radius: 16px;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.82);
    }
    .viz-topbar div[data-testid="stHorizontalBlock"] {
        gap: 0.55rem;
    }
    .viz-topbar [data-baseweb="select"] > div,
    .viz-topbar [data-baseweb="radio"] {
        background: rgba(255,255,255,0.78);
        border-radius: 12px;
    }
    .viz-main-card {
        padding: 0.95rem 1rem;
        margin-bottom: 0.75rem;
        border-radius: 18px;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(248,251,255,0.98) 100%);
        border: 1px solid #d7e4f4;
        box-shadow: 0 10px 28px rgba(15, 23, 42, 0.05);
    }
    .viz-main-card + .viz-main-card {
        margin-top: 0.1rem;
    }
    .viz-main-card div[data-testid="stHorizontalBlock"] {
        gap: 0.65rem;
    }
    .viz-panel-label {
        font-size: 0.86rem;
        font-weight: 700;
        margin-bottom: 0.38rem;
        letter-spacing: 0.01em;
    }
    .viz-panel-note {
        margin: 0 0 0.45rem 0;
        color: #60748c;
        font-size: 0.79rem;
        line-height: 1.55;
    }
    .viz-secondary-card {
        padding: 0.82rem 0.9rem;
        margin-bottom: 0.72rem;
        border-radius: 16px;
        background: linear-gradient(180deg, #fbfdff 0%, #f7fbff 100%);
        border: 1px solid #e1ebf7;
    }
    .viz-secondary-card div[data-testid="stExpander"] {
        border: 1px solid #e4edf8;
        border-radius: 14px;
        background: rgba(255,255,255,0.74);
        margin-bottom: 0.52rem;
        overflow: hidden;
    }
    .viz-secondary-card details {
        background: transparent;
    }
    .viz-secondary-card summary {
        font-weight: 600;
    }
    .viz-secondary-card div[data-testid="stExpanderDetails"] {
        padding-top: 0.18rem;
    }
    .viz-compact-status {
        padding: 0.62rem 0.76rem;
        margin: 0.55rem 0 0.75rem 0;
        border-radius: 12px;
        background: linear-gradient(180deg, #f4f8ff 0%, #eef6ff 100%);
        border: 1px solid #d7e7fb;
        color: #3d5b80;
        font-size: 0.8rem;
    }
    .viz-status-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem;
        margin: 0 0 0.7rem 0;
    }
    .viz-status-pill {
        display: inline-flex;
        align-items: center;
        gap: 0.28rem;
        padding: 0.3rem 0.66rem;
        border-radius: 999px;
        background: linear-gradient(180deg, #f3f8ff 0%, #edf5ff 100%);
        border: 1px solid #d7e7fb;
        color: #33567e;
        font-size: 0.76rem;
        line-height: 1;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.82);
    }
    .viz-status-pill strong {
        color: #11345d;
        font-weight: 700;
    }
    .viz-quiet-block {
        margin-top: 0.35rem;
        color: #6a7c92;
        font-size: 0.78rem;
        line-height: 1.5;
    }
    .viz-preview-shell {
        width: 100%;
        border: 1px solid #dbe7f5;
        border-radius: 18px;
        background: linear-gradient(180deg, #ffffff 0%, #f6faff 100%);
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.9), 0 10px 28px rgba(15, 23, 42, 0.05);
        padding: 0.9rem;
        overflow: hidden;
    }
    .viz-preview-frame {
        width: 100%;
        height: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 14px;
        background:
            radial-gradient(circle at top, rgba(191, 219, 254, 0.38), rgba(255,255,255,0.96)),
            linear-gradient(135deg, rgba(241,247,255,0.85) 0%, rgba(255,255,255,0.96) 100%);
        border: 1px dashed #c7d6ea;
    }
    .viz-preview-frame img {
        max-width: 100%;
        width: auto;
        height: auto;
        object-fit: contain;
        border-radius: 14px;
        box-shadow: 0 14px 30px rgba(15, 23, 42, 0.12);
    }
    .viz-preview-empty {
        max-width: 78%;
        text-align: center;
        color: #64748b;
        font-size: 0.92rem;
        line-height: 1.6;
    }
    .workbench-title {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 1rem;
        margin-bottom: 0.5rem;
    }
    .workbench-title h3 {
        margin: 0;
        color: #0f172a;
        font-size: 1.05rem;
    }
    .workbench-title p {
        margin: 0.2rem 0 0 0;
        color: #475569;
        font-size: 0.86rem;
    }
    .workbench-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.22rem 0.58rem;
        border-radius: 999px;
        background: #dbeafe;
        color: #1d4ed8;
        font-size: 0.76rem;
        font-weight: 600;
        white-space: nowrap;
    }
    .dashboard-stat {
        padding: 0.78rem 0.9rem;
        min-height: 104px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .dashboard-stat-label {
        font-size: 0.74rem;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #64748b;
        margin-bottom: 0.4rem;
    }
    .dashboard-stat-value {
        font-size: 1rem;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.3;
    }
    .dashboard-stat-meta {
        font-size: 0.8rem;
        color: #475569;
        margin-top: 0.3rem;
    }
    .subtle-kpi-row {
        padding: 0.2rem 0 0.55rem 0;
        color: #475569;
        font-size: 0.84rem;
    }
    .subtle-kpi-row strong {
        color: #0f172a;
    }
    .lab-thin-header {
        padding: 0.74rem 0.95rem;
        margin-bottom: 0.7rem;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(245,249,255,0.98) 100%);
        border: 1px solid #dbe7f5;
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
    }
    .lab-thin-header-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
    }
    .lab-thin-header h3 {
        margin: 0;
        font-size: 1.02rem;
        color: #10233f;
    }
    .lab-thin-header p {
        margin: 0.16rem 0 0 0;
        font-size: 0.83rem;
        color: #5a6f89;
    }
    .lab-main-card {
        padding: 0.95rem 1rem;
        margin-bottom: 0.75rem;
        border-radius: 18px;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(248,251,255,0.98) 100%);
        border: 1px solid #d7e4f4;
        box-shadow: 0 10px 28px rgba(15, 23, 42, 0.05);
    }
    .lab-secondary-card {
        padding: 0.82rem 0.9rem;
        margin-bottom: 0.72rem;
        border-radius: 16px;
        background: linear-gradient(180deg, #fbfdff 0%, #f7fbff 100%);
        border: 1px solid #e1ebf7;
    }
    .lab-panel-label {
        font-size: 0.86rem;
        font-weight: 700;
        margin-bottom: 0.38rem;
        color: #10233f;
        letter-spacing: 0.01em;
    }
    .lab-panel-note {
        margin: 0 0 0.45rem 0;
        color: #60748c;
        font-size: 0.79rem;
        line-height: 1.55;
    }
    .lab-status-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem;
        margin: 0 0 0.7rem 0;
    }
    .lab-status-pill {
        display: inline-flex;
        align-items: center;
        gap: 0.28rem;
        padding: 0.3rem 0.66rem;
        border-radius: 999px;
        background: linear-gradient(180deg, #f3f8ff 0%, #edf5ff 100%);
        border: 1px solid #d7e7fb;
        color: #33567e;
        font-size: 0.76rem;
        line-height: 1;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.82);
    }
    .lab-status-pill strong {
        color: #11345d;
        font-weight: 700;
    }
    .lab-quiet-note {
        color: #6a7c92;
        font-size: 0.78rem;
        line-height: 1.5;
    }
    .lab-secondary-card div[data-testid="stExpander"] {
        border: 1px solid #e4edf8;
        border-radius: 14px;
        background: rgba(255,255,255,0.74);
        margin-bottom: 0.52rem;
        overflow: hidden;
    }
    .lab-secondary-card summary {
        font-weight: 600;
    }
    .result-toolbar {
        margin: 0.55rem 0 0.15rem 0;
        padding: 0.68rem 0.8rem;
        background: linear-gradient(180deg, #f8fbff 0%, #f4f8ff 100%);
        border: 1px solid #dce8f6;
        border-radius: 12px;
    }
    .sidebar-section-note {
        font-size: 0.8rem;
        color: #64748b;
        margin-bottom: 0.2rem;
    }
    div.stButton > button, div.stDownloadButton > button {
        margin: 0.12rem 0.28rem 0.45rem 0;
        border-radius: 10px;
    }
    div.stDownloadButton > button[kind=\"secondary\"] {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        border: 1px solid #1d4ed8 !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.22);
        font-weight: 600;
    }
    div.stDownloadButton > button[kind=\"secondary\"]:hover {
        border-color: #1e40af !important;
        background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
        color: #ffffff !important;
    }
    div[data-testid=\"stTextArea\"] {
        margin-top: 0 !important;
        padding-top: 0 !important;
    }
    div[data-testid=\"stTextArea\"] > label {
        margin-top: 0 !important;
        padding-top: 0 !important;
        min-height: 0 !important;
    }
    .stTextArea textarea {
        font-family: \"Iosevka Term\", \"Sarasa Mono SC\", \"Noto Serif CJK SC\", \"Source Han Serif SC\", serif;
        line-height: 1.72;
        font-size: 0.96rem;
    }
    .stTextArea [data-baseweb=\"textarea\"] {
        border-radius: 12px;
        background: #fcfdff;
        border: 1px solid #dbe7f5;
    }
    .stTextArea [data-baseweb=\"textarea\"]:focus-within {
        border-color: #93c5fd;
        box-shadow: 0 0 0 1px #93c5fd;
    }
    .compact-caption {
        color: #64748b;
        font-size: 0.8rem;
        margin-top: -0.15rem;
        margin-bottom: 0.45rem;
    }
    .format-preview-box {
        min-height: 340px;
    }
    .format-chat-anchor {
        position: sticky;
        bottom: 0.75rem;
        z-index: 20;
        padding-top: 0.75rem;
        background: linear-gradient(180deg, rgba(243,246,251,0) 0%, rgba(243,246,251,0.92) 35%, rgba(243,246,251,1) 100%);
    }
</style>
""",
    },
}

SECTION_NAMES = list(CONFIG["sections"].keys())
FUNCTION_MATRIX = CONFIG["functions"]
FUNCTION_GROUPS = CONFIG["function_groups"]
FUNCTION_NAV = CONFIG["function_nav"]
DOMAIN_PROFILES = CONFIG["domains"]
DOMAIN_ORDER = list(DOMAIN_PROFILES.keys())
SECTIONS = CONFIG["sections"]
GLOBAL_HARD_LOCK_REGEX = CONFIG["global_hard_lock_regex"]
BUILTIN_LOCAL_SKILLS = CONFIG["builtin_local_skills"]
UI_CONFIG = CONFIG["ui"]
MODIFICATION_FUNCTIONS = {"馃搵 Redlining淇", "鉁?琛ㄨ揪娑﹁壊", "馃З 鏂囨湰闄嶉噸", "馃 鍘籄I鍛?(Humanizer)", "馃幆 绮句慨妯″紡"}
SHADOW_FUNCTIONS = {"鉁嶏笍 閫愭璧疯崏", "馃挕 鐮旂┒鎯虫硶鏋勬€?, "馃搫 鑺傝妭澶磋剳椋庢毚", "鉁嶏笍 褰卞瓙鍐欎綔"}
VIZ_SCENE_PROMPTS = {
    "鏈虹悊绀烘剰鍥?: "show the core scientific mechanism, structural relationships, interaction pathways, key local zoom-ins, and cause-effect logic",
    "缁撴瀯琛ㄥ緛鍥?: "focus on morphology, layered architecture, porosity, interfaces, crystallographic or nanoscale structural features",
    "鍚堟垚娴佺▼鍥?: "show a step-by-step synthesis or fabrication workflow with clear transitions between precursors, intermediates, and final products",
    "鎬ц兘瀵规瘮鍥捐В": "highlight comparative advantages, structure-property relationships, and visual evidence supporting performance claims",
    "鐣岄潰鍙嶅簲鍥?: "emphasize interface structure, interfacial reactions, boundary layers, and coupled transport or conversion processes",
    "浼犺緭璺緞鍥?: "emphasize ion, electron, heat, mass, or charge transport pathways with directional clarity",
    "瀹為獙娴佺▼鍥?: "show instruments, process stages, sample preparation, testing order, and workflow logic clearly",
    "閫昏緫妗嗘灦鍥?: "present conceptual nodes, relationships, hierarchy, and scientific reasoning in a structured visual map",
}
VIZ_STYLE_PROMPTS = {
    "绉戠爺 3D 娓叉煋": "scientific 3D rendering, realistic material texture, layered depth, polished academic illustration",
    "BioRender 椋庢牸": "clean biomedical-style schematic, crisp icons, simplified but professional scientific composition",
    "Nature 鍥惧舰鎽樿椋庢牸": "high-end journal graphical abstract style, concise layout, premium composition, strong clarity",
    "绠€绾︾煝閲忛鏍?: "minimal vector scientific illustration, clean edges, reduced clutter, publication-ready simplicity",
    "鎵佸钩鍖栦俊鎭浘椋庢牸": "flat infographic style, clear hierarchy, simplified geometry, explanatory visual balance",
    "娣辫壊楂樼骇鎰熼鏍?: "dark premium scientific visual style, cinematic contrast, glowing highlights, elegant composition",
    "楂樺姣旀紨绀洪鏍?: "presentation-oriented, high contrast, visually striking, immediately readable on slides",
}
VIZ_DEFAULT_STATE = {
    "material_name": "",
    "component_tags_text": "",
    "component_tags": [],
    "usage": "璁烘枃涓诲浘",
    "scene": "鏈虹悊绀烘剰鍥?,
    "style": "绉戠爺 3D 娓叉煋",
    "emphasis_points_text": "",
    "structure_notes": "",
    "description": "",
    "label_mode": "鏃犳枃瀛楃増",
    "label_language": "涓枃",
    "label_content_mode": "鑷姩鐢熸垚鍚庣紪杈戯紙鎺ㄨ崘锛?,
    "label_terms_text": "",
    "auto_generated_labels": [],
    "final_label_terms": [],
    "final_language_rule": "绂佹浠讳綍鏂囧瓧鏍囩",
    "info_density": "涓?,
    "aspect_ratio": "1:1",
    "logic_summary": "",
    "skill_prompt": "",
    "optimized_prompt": "",
    "hard_constraints": "",
    "base_prompt": "",
    "prompt_with_labels": "",
    "prompt_without_labels": "",
    "compact_prompt": "",
    "expanded_prompt": "",
    "final_prompt": "",
    "current_image_url": "",
    "current_image_bytes": b"",
    "current_result_id": None,
    "current_parent_id": None,
    "iteration_instruction": "",
    "local_area_hint": "",
    "inpaint_strength": 0.35,
    "preserve_composition": True,
    "edit_mode": "灞€閮ㄨ繘鍖?,
    "iteration_mode": "text-to-image",
    "seed_value": None,
    "current_seed": None,
    "current_image_seed": None,
    "style_reference_image": None,
    "style_reference_name": "",
    "style_strength": 0.6,
    "history": [],
    "active_history_id": None,
    "generation_counter": 0,
    "is_generating": False,
    "last_error": "",
}
PPT_PAGE_TYPES = ["灏侀潰椤?, "鑳屾櫙椤?, "闂瀹氫箟椤?, "鏂规硶椤?, "娴佺▼椤?, "缁撴灉椤?, "瀵规瘮椤?, "鏈虹悊椤?, "缁撹椤?, "灞曟湜椤?]
PPT_STYLE_MODES = ["academic-paperskills", "journal-briefing", "defense-clean"]
PPT_DEFAULT_LAYOUTS = {
    "宸︽枃鍙冲浘": {"title": "涓婃柟鏍囬 + 宸︽枃鍙冲浘", "hint": "閫傚悎鑳屾櫙銆佹柟娉曘€佺粨鏋滆鏄?},
    "涓婂浘涓嬫枃": {"title": "涓婂浘涓嬫枃", "hint": "閫傚悎娴佺▼銆佺粨鏋滃睍绀恒€佺粨鏋勮鏄?},
    "鍙屾爮瀵规瘮": {"title": "鍙屾爮瀵规瘮", "hint": "閫傚悎 before/after銆佹ā鍨嬪姣斻€佸疄楠屽姣?},
    "澶у浘閲嶇偣璇存槑": {"title": "澶у浘閲嶇偣璇存槑", "hint": "閫傚悎鍗曞浘寮鸿皟銆佹満鐞嗗浘閰嶈鏄?},
    "缁撴灉+缁撹": {"title": "缁撴灉 + 缁撹", "hint": "閫傚悎瀹為獙缁撴灉涓庝竴鍙ヨ瘽缁撹骞剁疆"},
    "鏈虹悊鍥捐鏄庨〉": {"title": "鏈虹悊鍥捐鏄庨〉", "hint": "閫傚悎绉戠爺缁樺浘涓庡叧閿?bullet 閰嶅悎"},
}
PPT_PAGE_TYPE_LAYOUTS = {
    "灏侀潰椤?: "澶у浘閲嶇偣璇存槑",
    "鑳屾櫙椤?: "宸︽枃鍙冲浘",
    "闂瀹氫箟椤?: "宸︽枃鍙冲浘",
    "鏂规硶椤?: "宸︽枃鍙冲浘",
    "娴佺▼椤?: "涓婂浘涓嬫枃",
    "缁撴灉椤?: "缁撴灉+缁撹",
    "瀵规瘮椤?: "鍙屾爮瀵规瘮",
    "鏈虹悊椤?: "鏈虹悊鍥捐鏄庨〉",
    "缁撹椤?: "缁撴灉+缁撹",
    "灞曟湜椤?: "宸︽枃鍙冲浘",
}
PPT_DEFAULT_STATE = {
    "input_mode": "绾枃瀛?,
    "page_title": "",
    "page_type": "缁撴灉椤?,
    "page_goal": "",
    "raw_text": "",
    "extra_notes": "",
    "uploaded_source_name": "",
    "uploaded_source_text": "",
    "uploaded_images": [],
    "reference_images": [],
    "template_name": "榛樿绉戠爺妯℃澘",
    "template_analysis": {},
    "template_file_name": "",
    "style_mode": "academic-paperskills",
    "text_simplify_level": "涓?,
    "info_density": "涓?,
    "keep_original_images": True,
    "allow_external_support": False,
    "generate_aux_figure": False,
    "apply_template_layout": True,
    "processed_text_variants": {},
    "selected_layout": "缁撴灉+缁撹",
    "layout_analysis": {},
    "current_slide_markdown": "",
    "current_slide_struct": {},
    "current_slide_preview_html": "",
    "current_slide_snapshot": {},
    "current_slide_id": None,
    "current_parent_id": None,
    "micro_tune_request": "",
    "assembly_pages": [],
    "history": [],
    "active_history_id": None,
    "generation_counter": 0,
    "last_error": "",
}
>>>>>>> e508ee5 (feat: switch default image model to imagen-2.0, fix sdk_version in README_HF)

# 鈹€鈹€ Session State 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€
DEFAULT_STATES = {
    "history": [],              # 鍘嗗彶鏃跺厜鏈猴紙鎸夋澘鍧?鍔熻兘鍒嗙被锛?
    "reference_docs": {},       # 鏍囨潌鏂囩尞搴?{filename: style_analysis}
    "current_input": "",        # 褰撳墠杈撳叆
    "loaded_skills": False,     # Skills鍔犺浇鐘舵€?
}

for key, default in DEFAULT_STATES.items():
    if key not in st.session_state:
        st.session_state[key] = default

# 鈹€鈹€ 甯搁噺瀹氫箟 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

# 15+ 鍔熻兘鐭╅樀锛堝崟涓€鑱岃矗锛?
FUNCTION_MATRIX = {
    "馃摑 涓浆鑻辩炕璇?: {
        "description": "浠呮墽琛孋N鈫扙N璇杞崲锛岀姝㈡鼎鑹?,
        "rules": ["鉂?绂佹鑷娑﹁壊", "鉂?绂佹鏀瑰彉鍘熸剰", "鉁?浠呰瑷€杞崲"]
    },
    "鉁?琛ㄨ揪娑﹁壊": {
        "description": "鎻愬崌瀛︽湳鍦伴亾鎬э紝鍚岃瑷€浼樺寲",
        "rules": ["鉂?绂佹缈昏瘧", "鉂?绂佹鏀瑰彉鍘熸剰", "鉁?浠呴檺鍚岃瑷€"]
    },
    "馃攳 閫昏緫妫€鏌?: {
        "description": "妫€鏌ュ洜鏋滈摼鏉″拰琛旀帴",
        "rules": ["鉂?浠呮煡閫昏緫", "鉂?涓嶄慨鏀规枃鏈?, "鉁?鎸囧嚭闂"]
    },
    "馃 鍘籄I鍛?(Humanizer)": {
        "description": "娑堥櫎AI鐥曡抗锛屾ā浠跨湡浜鸿搴?,
        "rules": ["鉂?绂佹缈昏瘧", "鉂?绂佹澧炲姞璁虹偣", "鉁?CN浼樺寲CN/EN浼樺寲EN"]
    },
    "馃懆鈥嶁殩锔?Reviewer瑙嗚": {
        "description": "妯℃嫙瀹＄浜鸿瑙掑瑙?,
        "rules": ["鉁?鎻愪緵鏀硅繘寤鸿", "鉁?鎸囧嚭钖勫急鐜妭"]
    },
    "馃挕 鐮旂┒鎯虫硶鏋勬€?: {
        "description": "浠庨浂鏋勬€濈爺绌舵柟鍚?,
        "rules": ["鉁?鍩轰簬鐮旂┒棰嗗煙", "鉁?鎻愪緵鍒涙柊鐐?]
    },
    "馃 ML璁烘枃鍐欎綔": {
        "description": "NeurIPS/ICML绾у埆鍐欎綔",
        "rules": ["鉁?寮曠敤鏍煎紡楠岃瘉", "鉁?鍥捐〃鎻忚堪瑙勮寖"]
    },
    "馃搳 姒傚康鍥捐璁?: {
        "description": "鐢熸垚璁烘枃姒傚康鍥捐璁?,
        "rules": ["鉁?璁捐鍝插", "鉁?缁樺浘鎻愮ず璇?]
    },
    "馃搫 鑺傝妭澶磋剳椋庢毚": {
        "description": "鎸夌珷鑺傝繘琛屽ご鑴戦鏆?,
        "rules": ["鉁?閽堝褰撳墠鏉垮潡", "鉁?鎻愪緵鎬濊矾"]
    },
    "鉁嶏笍 閫愭璧疯崏": {
        "description": "灏嗗ぇ绾叉墿鍏呬负姝ｅ紡娈佃惤",
        "rules": ["鉁?妯′豢鏍囨潌鏂囩尞", "鉁?淇濇寔瀛︽湳瑙勮寖"]
    },
    "馃幆 绮句慨妯″紡": {
        "description": "娣卞害绮句慨锛岃揪鍒伴《鍒婃按鍑?,
        "rules": ["鉁?妯′豢鏍囨潌鏂囩尞", "鉁?椤跺垔鏍囧噯"]
    },
    "馃摑 寮曠敤楠岃瘉": {
        "description": "妫€鏌ュ紩鐢ㄦ牸寮忓拰瀹屾暣鎬?,
        "rules": ["鉁?鏍煎紡楠岃瘉", "鉁?瀹屾暣鎬ф鏌?]
    },
    "馃帹 鍥捐〃瑙勮寖妫€鏌?: {
        "description": "妫€鏌ュ浘琛ㄦ弿杩拌鑼冩€?,
        "rules": ["鉁?鑹茬洸鍙嬪ソ", "鉁?瑙勮寖楠岃瘉"]
    },
    "馃搵 Redlining淇": {
        "description": "甯︿慨璁㈢棔杩圭殑淇敼寤鸿",
        "rules": ["鉁?鏄剧ず淇敼鐥曡抗", "鉁?瀵规瘮瑙嗗浘"]
    },
    "馃攧 鐗堟湰瀵规瘮": {
        "description": "瀵规瘮涓嶅悓鐗堟湰宸紓",
        "rules": ["鉁?楂樹寒宸紓", "鉁?淇敼璇存槑"]
    },
}

# 鏉垮潡瀹氫箟
SECTIONS = {
    "鎽樿": {"focus": "寮€闂ㄨ灞憋紝鏁版嵁鏀拺", "max_words": 250},
    "寮曡█": {"focus": "鑳屾櫙杞姌锛岀爺绌剁┖鐧?, "max_words": 800},
    "鏂规硶": {"focus": "娴佺▼娓呮櫚锛屽弬鏁扮簿纭?, "max_words": 1500},
    "缁撴灉": {"focus": "鏁版嵁瀹㈣锛屽浘琛ㄥ紩鐢?, "max_words": 1200},
    "璁ㄨ": {"focus": "娣卞害瑙ｈ锛屾枃鐚姣?, "max_words": 1000},
    "缁撹": {"focus": "鎬荤粨璐＄尞锛屽睍鏈涙湭鏉?, "max_words": 300},
}

# 瀛︾棰嗗煙涓庢湳璇‖閿?
DOMAINS = {
    "馃攱 鑳芥簮鐢垫睜": {
        "hard_lock": [r'\$Li\+\$', r'\$Na\+\$', 'NCM523', 'NCM622', 'capacity retention', 'intercalation'],
        "focus": "鐢靛寲瀛︽€ц兘"
    },
    "馃彈锔?鍥哄簾/鍦熸湪": {
        "hard_lock": ['GGBS', '姘村寲鍔ㄥ姏瀛?, 'compressive strength', 'C-S-H', 'pozzolanic'],
        "focus": "鏉愭枡鎬ц兘"
    },
    "馃敥 鏈烘/鏉愭枡": {
        "hard_lock": ['tensile strength', 'grain boundary', 'dislocation', 'microstructure'],
        "focus": "鍔涘鎬ц兘"
    },
    "馃И 鍌寲/鍖栧伐": {
        "hard_lock": ['TOF', 'turnover frequency', 'BET', 'heterogeneous catalysis'],
        "focus": "鍌寲鎬ц兘"
    },
    "馃寠 鐜宸ョ▼": {
        "hard_lock": ['COD', 'BOD', 'MBR', 'activated sludge', 'removal efficiency'],
        "focus": "澶勭悊鏁堟灉"
    },
    "馃摫 鐢靛瓙鍗婂浣?: {
        "hard_lock": ['MOSFET', 'bandgap', 'GaN', 'SiC', 'carrier mobility'],
        "focus": "鐢靛鎬ц兘"
    },
    "馃К 鐢熺墿鏉愭枡": {
        "hard_lock": ['hydrogel', 'cell adhesion', 'biocompatibility', 'MTT assay'],
        "focus": "鐢熺墿鐩稿鎬?
    },
    "馃 鏈哄櫒瀛︿範": {
        "hard_lock": ['CNN', 'Transformer', 'overfitting', 'gradient descent', 'F1-score'],
        "focus": "妯″瀷鎬ц兘"
    },
    "馃捇 杞欢宸ョ▼": {
        "hard_lock": ['CI/CD', 'Agile', 'DevOps', 'API', 'microservices'],
        "focus": "宸ョ▼瀹炶返"
    },
}

# 鈹€鈹€ 娣卞害鎶€鑳藉姞杞界郴缁?鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

@st.cache_resource
def load_local_skills() -> Dict[str, str]:
    """閫掑綊鍔犺浇鏈湴瀛︽湳搴?""
    skills = {}
    base_path = Path("./awesome-ai-research-writing")

    if not base_path.exists():
        return skills

    for md_file in base_path.rglob("*.md"):
        try:
            content = md_file.read_text(encoding="utf-8")
            key = md_file.stem
            skills[key] = content
        except:
            continue

    return skills

LOCAL_SKILLS = load_local_skills()

# Humanizer 鏍稿績瑙勫垯锛堢‖缂栫爜娉ㄥ叆锛?
HUMANIZER_RULES = """
## Humanizer 鏍稿績瑙勫垯锛堥浂绡′綅锛?

### AI 鐥曡抗璇嗗埆娓呭崟
1. **杩囧害寮鸿皟鎰忎箟**: "鍏锋湁...鎰忎箟", "crucial", "critical", "important", "significant"
2. **AI 甯哥敤璇?*: "delve", "landscape", "realm", "leverage", "underscore", "utilize"
3. **鐮存姌鍙锋互鐢?*: 瑙ｉ噴鎬ф彃鍏ヨ杩囧锛堚€攚hich, 鈥攖hat锛?
4. **涓夌偣寮忓爢鐮?*: A, B, and C 缁撴瀯杩囧
5. **淇冮攢鑵?*: "exciting", "promising", "novel", "groundbreaking"
6. **绌烘礊-ing鍒嗘瀽**: "indicating", "suggesting", "implying" 婊ョ敤

### 浜哄懗娉ㄥ叆绛栫暐
- 鎵胯涓嶇‘瀹氭€? "suggest", "may", "potentially", "appears to"
- 鑺傚鍙樺寲: 闀跨煭鍙ヤ氦鏇匡紝閬垮厤鍗曡皟
- 鑷劧杩囨浮: 鍒犻櫎鏈烘杩炴帴璇嶏紙Firstly, Secondly, Furthermore锛?
- 绠€娲佸姩璇? show, use, find, make (鑰岄潪 demonstrate, utilize, discover, fabricate)

### 璇簭浼樺寲鍘熷垯
- **涓枃杈撳叆**: 浼樺寲涓枃璇簭锛屼繚鎸佷腑鏂囪緭鍑?
- **鑻辨枃杈撳叆**: 澧炲己鑻辨枃鑺傚鎰燂紝淇濇寔鑻辨枃杈撳嚭
- **闆剁炕璇?*: 涓ユ牸绂佹璇杞崲
"""

# ML Paper Writing Checklist锛堜粎鐢ㄤ簬 Results 鏉垮潡楠岃瘉锛?
ML_CHECKLIST = """
## ML Paper Writing Checklist (浠?Results 鏉垮潡)

### 寮曠敤楠岃瘉
- [ ] 寮曠敤鏍煎紡缁熶竴锛圢eurIPS/ICML 鏍煎紡锛?
- [ ] 鎵€鏈夋柟娉曞紩鐢ㄦ湁鏄庣‘鍑哄
- [ ] 鍩虹嚎妯″瀷姝ｇ‘寮曠敤

### 鍥捐〃瑙勮寖
- [ ] 鑹茬洸鍙嬪ソ閰嶈壊锛坴iridis, magma锛?
- [ ] 鍧愭爣杞存爣绛炬竻鏅?
- [ ] 鍥句緥浣嶇疆鍚堢悊
- [ ] 璇樊绾挎爣娉?

### 鏁版嵁瀹屾暣鎬?
- [ ] 瓒呭弬鏁拌〃瀹屾暣
- [ ] 瀹為獙璁剧疆鍙鐜?
- [ ] 缁熻鏄捐憲鎬ф爣娉?
"""


# 鈹€鈹€ 鏈纭攣淇濇姢绯荤粺 v2.0 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

def protect_hard_terms(text: str, domain: str) -> Tuple[str, Dict[str, str]]:
    """纭攣淇濇姢宸ョ鏈锛岀粷瀵圭姝㈡敼鍔?v2.0"""
    protected = text
    mapping = {}
    counter = 0

    # 1. 淇濇姢 LaTeX 鍏紡锛堝妯″紡锛?
    # 琛屽唴鍏紡: $...$
    for match in re.finditer(r'\$[^$]+?\$', text):
        placeholder = f"__LOCK_{counter}__"
        mapping[placeholder] = match.group()
        protected = protected.replace(match.group(), placeholder, 1)
        counter += 1

    # 鍧楃骇鍏紡: $$...$$ 鎴?\[...\]
    for match in re.finditer(r'\$\$[^$]+?\$\$|\\\[.*?\\\]', text, re.DOTALL):
        placeholder = f"__LOCK_{counter}__"
        mapping[placeholder] = match.group()
        protected = protected.replace(match.group(), placeholder, 1)
        counter += 1

    # LaTeX鍛戒护: \frac{}{}, \sum_{}^{}, \int_{}^{}, etc.
    for match in re.finditer(r'\\[a-zA-Z]+(?:\{[^}]*\}|\[[^\]]*\]|_[^{}s]|[_^]{[^}]*})?', text):
        placeholder = f"__LOCK_{counter}__"
        mapping[placeholder] = match.group()
        protected = protected.replace(match.group(), placeholder, 1)
        counter += 1

    # 2. 淇濇姢鍖栧寮忓拰鏉愭枡鍚嶇О
    chemical_patterns = [
        r'\b[A-Z][a-z]?\d*(?:_[\d-]+|[+\-]?\d*)?\b',  # Li+, Na+, NCM523
        r'\b[A-Z][a-z]?\d*(?:_[\d-]+|[+\-]?\d*)?\s*[A-Z][a-z]?\d*',  # LiCoO2, NaMnO2
    ]
    for pattern in chemical_patterns:
        for match in re.finditer(pattern, text):
            # 鎺掗櫎鏅€氬崟璇?
            if len(match.group()) > 2 and any(c.isdigit() or c in '_+-' for c in match.group()):
                placeholder = f"__LOCK_{counter}__"
                mapping[placeholder] = match.group()
                protected = protected.replace(match.group(), placeholder, 1)
                counter += 1

    # 3. 淇濇姢棰嗗煙鐗瑰畾鏈
    if domain in DOMAINS:
        for term in DOMAINS[domain]["hard_lock"]:
            if isinstance(term, str):
                # 绮剧‘鍖归厤
                if term in protected:
                    placeholder = f"__LOCK_{counter}__"
                    mapping[placeholder] = term
                    protected = protected.replace(term, placeholder, 1)
                    counter += 1

    # 4. 淇濇姢娴嬮噺鍗曚綅鍜屾暟鍊肩粍鍚?
    unit_patterns = [
        r'\d+\s*(?:mAh路g[^-1]|mAh g[^-1]|mA h[^-1]|A h[^-1])',
        r'\d+\s*(?:掳C|K|MPa|GPa|kPa|Pa)',
        r'\d+\s*(?:wt%|vol%|at%|mol%)',
        r'\d+\s*(?:nm|渭m|mm|cm|m|km)',
        r'\d+\s*(?:Hz|kHz|MHz|GHz)',
        r'\d+\s*(?:S cm[^-1]|S m[^-1])',
    ]
    for pattern in unit_patterns:
        for match in re.finditer(pattern, text):
            placeholder = f"__LOCK_{counter}__"
            mapping[placeholder] = match.group()
            protected = protected.replace(match.group(), placeholder, 1)
            counter += 1

    return protected, mapping


def restore_hard_terms(text: str, mapping: Dict[str, str]) -> str:
    """鎭㈠纭攣淇濇姢鐨勬湳璇?""
    result = text
    for placeholder, original in mapping.items():
        result = result.replace(placeholder, original)
    return result


# 鈹€鈹€ 璇█鏅鸿兘璇嗗埆 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

def detect_language(text: str) -> str:
    """鏅鸿兘璇嗗埆杈撳叆璇█"""
    # 缁熻涓枃瀛楃
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
    total_chars = len(text.strip())

    if total_chars == 0:
        return "unknown"

    chinese_ratio = chinese_chars / total_chars

    if chinese_ratio > 0.3:
        return "zh"
    else:
        return "en"


# 鈹€鈹€ 褰卞瓙鍚堣憲鑰咃細鏍囨潌鏂囩尞鍒嗘瀽 v2.0 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

def extract_language_genes(text: str) -> Dict:
    """娣卞害鎻愬彇鏂囩尞鐨勮瑷€鍩哄洜"""
    genes = {
        "sentence_length_pattern": [],
        "connecting_words": Counter(),
        "academic_phrases": Counter(),
        "voice_pattern": {"passive": 0, "active": 0},
        "citation_style": Counter(),
        "complexity_markers": Counter(),
    }

    sentences = [s.strip() for s in text.split('.') if s.strip()]
    for sent in sentences[:100]:  # 鍒嗘瀽鍓?00鍙?
        # 鍙ラ暱妯″紡
        words = len(sent.split())
        genes["sentence_length_pattern"].append(words)

        # 杩炴帴璇嶇粺璁?
        connectors = [
            'however', 'therefore', 'furthermore', 'moreover', 'additionally',
            'consequently', 'subsequently', 'meanwhile', 'nevertheless',
            'thus', 'hence', 'accordingly', 'otherwise', 'moreover'
        ]
        for conn in connectors:
            if re.search(rf'\b{conn}\b', sent.lower()):
                genes["connecting_words"][conn] += 1

        # 瀛︽湳鐭
        phrases = [
            'it is worth noting', 'it should be mentioned', 'previous studies',
            'recent work', 'to the best of', 'knowledge', 'suggest that',
            'indicate that', 'demonstrate that', 'reveal that'
        ]
        for phrase in phrases:
            if phrase in sent.lower():
                genes["academic_phrases"][phrase] += 1

        # 涓诲姩/琚姩璇€?
        if re.search(r'\b(was|were)\s+\w+ed\b', sent, re.IGNORECASE):
            genes["voice_pattern"]["passive"] += 1
        else:
            genes["voice_pattern"]["active"] += 1

        # 寮曠敤椋庢牸
        citations = re.findall(r'\[\d+\]|\([^)]+\d+[^)]*\)|\w+\s+et\s+al\.', sent)
        for cite in citations:
            genes["citation_style"][cite[:20]] += 1

        # 澶嶆潅鎬ф爣璁?
        complexity = [
            'although', 'while', 'despite', 'whereas', 'not only but also',
            'either or', 'neither nor', 'whether or'
        ]
        for marker in complexity:
            if re.search(rf'\b{marker}\b', sent.lower()):
                genes["complexity_markers"][marker] += 1

    # 璁＄畻缁熻鏁版嵁
    if genes["sentence_length_pattern"]:
        genes["avg_sentence_length"] = sum(genes["sentence_length_pattern"]) / len(genes["sentence_length_pattern"])
        genes["sentence_length_std"] = (sum((x - genes["avg_sentence_length"])**2 for x in genes["sentence_length_pattern"]) /
                                       len(genes["sentence_length_pattern"]))**0.5
    else:
        genes["avg_sentence_length"] = 0
        genes["sentence_length_std"] = 0

    return genes


def analyze_reference_paper(docx_bytes: bytes, filename: str) -> str:
    """鍒嗘瀽鏍囨潌鏂囩尞鐨勫啓浣滈鏍?v2.0 - 璇█鍩哄洜娣卞害鎻愬彇"""
    try:
        doc = Document(BytesIO(docx_bytes))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        text = "\n".join(paragraphs[:80])  # 鍙栧墠80娈?

        # 鎻愬彇椋庢牸鐗瑰緛
        lang = detect_language(text)
        genes = extract_language_genes(text)

        # 鐢熸垚椋庢牸鍒嗘瀽鎶ュ憡
        style_analysis = f"""
## 馃К 鏍囨潌鏂囩尞璇█鍩哄洜鎻愬彇: {filename}

### 馃搳 鍩虹淇℃伅
- **娈佃惤鏁?*: {len(paragraphs)}
- **璇█**: {"涓枃" if lang == "zh" else "鑻辨枃"}
- **骞冲潎鍙ラ暱**: {genes.get('avg_sentence_length', 0):.1f} 璇?鍙?
- **鍙ラ暱娉㈠姩**: {genes.get('sentence_length_std', 0):.1f} (鏍囧噯宸?

### 馃幍 鍙欎簨鑺傚
- **骞冲潎鍙ラ暱**: {genes.get('avg_sentence_length', 0):.1f} 璇?鍙?
- **鍙ュ紡澶氭牱鎬?*: {"楂? if genes.get('sentence_length_std', 0) > 10 else "涓? if genes.get('sentence_length_std', 0) > 5 else "浣?}

### 馃敆 杩炴帴璇嶅亸濂?
{', '.join([f'{k}({v})' for k, v in genes['connecting_words'].most_common(5)]) if genes['connecting_words'] else '鏃犳槑鏄惧亸濂?}

### 馃帗 瀛︽湳鐭
{', '.join([f'{k}({v})' for k, v in genes['academic_phrases'].most_common(5)]) if genes['academic_phrases'] else '鏃犲父鐢ㄧ煭璇?}

### 馃摙 璇€佸€惧悜
- **涓诲姩**: {genes['voice_pattern']['active']} 鍙?
- **琚姩**: {genes['voice_pattern']['passive']} 鍙?
- **涓诲姩姣?*: {genes['voice_pattern']['active'] / max(1, genes['voice_pattern']['active'] + genes['voice_pattern']['passive']) * 100:.1f}%

### 馃摎 寮曠敤椋庢牸
{', '.join([k for k, v in genes['citation_style'].most_common(3)]) if genes['citation_style'] else '鏈娴嬪埌鏄庣‘妯″紡'}

### 馃 澶嶆潅鎬ф爣璁?
{', '.join([f'{k}({v})' for k, v in genes['complexity_markers'].most_common(5)]) if genes['complexity_markers'] else '鏃犳槑鏄惧鏉傜粨鏋?}

### 馃摑 鏍锋湰鏂囨湰
{text[:500]}...

### 馃幆 妯′豢绛栫暐
1. **鍙ラ暱鎺у埗**: 鐩爣骞冲潎鍙ラ暱 {genes.get('avg_sentence_length', 0):.0f} 璇嶏紝娉㈠姩 卤{genes.get('sentence_length_std', 0):.0f}
2. **杩炴帴璇?*: 浼樺厛浣跨敤 {', '.join([k for k, v in genes['connecting_words'].most_common(3)]) if genes['connecting_words'] else '鑷劧杩囨浮'}
3. **璇€?*: {"鍋忓ソ涓诲姩璇€? if genes['voice_pattern']['active'] > genes['voice_pattern']['passive'] * 1.5 else "骞宠　涓诲姩/琚姩"}
4. **澶嶆潅鎬?*: {"楂樺鏉傚彞寮? if sum(genes['complexity_markers'].values()) > 5 else "涓瓑澶嶆潅搴?}
"""
        return style_analysis
    except Exception as e:
        return f"鍒嗘瀽澶辫触: {e}"


# 鈹€鈹€ 闆剁浣?Prompt 鐢熸垚鍣?鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

def create_strict_prompt(
    function: str,
    input_text: str,
    section: str,
    domain: str,
    reference_styles: List[str],
    lang: str
) -> str:
    """鍒涘缓鍗曚竴鑱岃矗 prompt锛屼弗鏍煎姛鑳介殧绂?v2.0"""

    # 璇█閿佸畾寮哄埗澹版槑
    lang_lock = {
        "zh": """## 鈿狅笍 璇█閿佸畾锛堝己鍒舵墽琛岋級
1. **妫€娴嬪埌涓枃杈撳叆**锛氬叏绋嬩娇鐢ㄤ腑鏂囧鐞?
2. **闆剁炕璇?*锛氫弗鏍肩姝换浣曚腑鈫掕嫳缈昏瘧琛屼负
3. **杈撳嚭璇█閿佸畾**锛氭墍鏈夎緭鍑哄繀椤绘槸涓枃
4. **鏈淇濇姢**锛氳嫳鏂囦笓鏈夊悕璇嶄繚鎸佸師鏍?"",
        "en": """## 鈿狅笍 Language Lock (Strict Enforcement)
1. **English Input Detected**: Process entirely in English
2. **Zero Translation**: Strictly forbidden from translating to any other language
3. **Output Language Lock**: All output must be in English
4. **Term Protection**: Preserve non-English technical terms as-is"""
    }

    # 鍩虹娉ㄥ叆
    base_injection = f"""
# XueYan OS v4.0 - 闆剁浣嶆墽琛岀郴缁?
## 馃敀 褰撳墠閰嶇疆
- **鍔熻兘妯″紡**: {function}
- **璁烘枃鏉垮潡**: {section}
- **鐮旂┒棰嗗煙**: {domain}
- **杈撳叆璇█**: {"涓枃" if lang == "zh" else "鑻辨枃"}
- **鏈淇濇姢**: 宸叉縺娲荤‖閿佷繚鎶?

## 馃摎 鏈湴瀛︽湳搴?
宸插姞杞?{len(LOCAL_SKILLS)} 涓鏈?Skills

{lang_lock.get(lang, '')}
"""

    # 娣诲姞 Humanizer 瑙勫垯
    if "Humanizer" in function or "鍘籄I鍛? in function:
        base_injection += f"""
## 馃幆 Humanizer 鏍稿績瑙勫垯锛坽lang.upper()}涓撶敤锛?
### AI 鐥曡抗璇嗗埆娓呭崟
1. **杩囧害寮鸿皟鎰忎箟**: "鍏锋湁...鎰忎箟", "crucial", "critical", "important", "significant"
2. **AI 甯哥敤璇?*: "delve", "landscape", "realm", "leverage", "underscore", "utilize"
3. **鐮存姌鍙锋互鐢?*: 瑙ｉ噴鎬ф彃鍏ヨ杩囧锛堚€攚hich, 鈥攖hat锛?
4. **涓夌偣寮忓爢鐮?*: A, B, and C 缁撴瀯杩囧
5. **淇冮攢鑵?*: "exciting", "promising", "novel", "groundbreaking"
6. **绌烘礊-ing鍒嗘瀽**: "indicating", "suggesting", "implying" 婊ョ敤

### 浜哄懗娉ㄥ叆绛栫暐
- **鎵胯涓嶇‘瀹氭€?*: "suggest", "may", "potentially", "appears to"
- **鑺傚鍙樺寲**: 闀跨煭鍙ヤ氦鏇匡紝閬垮厤鍗曡皟
- **鑷劧杩囨浮**: 鍒犻櫎鏈烘杩炴帴璇嶏紙Firstly, Secondly, Furthermore锛?
- **绠€娲佸姩璇?*: show, use, find, make (鑰岄潪 demonstrate, utilize, discover, fabricate)

### 璇簭浼樺寲鍘熷垯
- **褰撳墠璇█**: {"涓枃" if lang == "zh" else "鑻辨枃"}
- **浼樺寲鐩爣**: {"浼樺寲涓枃璇簭锛屼繚鎸佷腑鏂囪緭鍑? if lang == "zh" else "澧炲己鑻辨枃鑺傚鎰燂紝淇濇寔鑻辨枃杈撳嚭"}
- **闆剁炕璇?*: 涓ユ牸绂佹璇杞崲
"""
    else:
        base_injection += HUMANIZER_RULES + "\n"

    # 濡傛灉鏄?Results 鏉垮潡涓旀湁 ML 鍐欎綔闇€姹傦紝娣诲姞 checklist
    if section == "缁撴灉" and "ML" in domain:
        base_injection += ML_CHECKLIST + "\n"

    # 娣诲姞鏍囨潌鏂囩尞椋庢牸锛堝奖瀛愬悎钁楄€咃細浠呴檺璧疯崏/鏋勬€濈被鍔熻兘锛?
    shadow_functions = {"鉁嶏笍 閫愭璧疯崏", "馃挕 鐮旂┒鎯虫硶鏋勬€?, "馃搫 鑺傝妭澶磋剳椋庢毚"}
    if reference_styles and function in shadow_functions:
        base_injection += "\n## 馃摎 褰卞瓙鍚堣憲鑰咃細鏍囨潌鏂囩尞璇█鍩哄洜\n"
        base_injection += "\n".join(reference_styles)
        base_injection += "\n\n**妯′豢鎸囦护**: 涓ユ牸妯′豢涓婅堪鏍囨潌鏂囩尞鐨勫彞闀裤€佽繛鎺ヨ瘝銆佽鎬佸拰澶嶆潅鎬у亸濂姐€俓n"

    # 鏉垮潡鐗瑰畾鎸囧
    section_info = SECTIONS.get(section, {})
    if section_info:
        base_injection += f"""
## 馃摑 {section} 鏉垮潡鍐欎綔瑕佹眰
- **鏍稿績閲嶇偣**: {section_info.get('focus', '')}
- **寤鸿闀垮害**: {section_info.get('max_words', '')} 璇?
- **鍐欎綔绛栫暐**: 閽堝璇ユ澘鍧楃殑鍙欎簨閫昏緫杩涜浼樺寲
"""

    # 鍔熻兘鐗瑰畾瑙勫垯锛堥浂绡′綅锛?
    func_rules = FUNCTION_MATRIX.get(function, {})
    if func_rules:
        base_injection += f"""
## 馃幆 {function} 鍔熻兘瑙勫垯锛堝崟涓€鑱岃矗锛?
{chr(10).join(f'- {rule}' for rule in func_rules.get('rules', []))}

**鉀?闆剁浣嶅己鍒跺０鏄?*: 涓ユ牸鎸夌収涓婅堪瑙勫垯鎵ц锛岀粷涓嶈秺鐣屽埌鍏朵粬鍔熻兘棰嗗煙銆?
"""

    # 鍔熻兘鐗瑰畾妯℃澘
    function_templates = {
        "馃摑 涓浆鑻辩炕璇?: f"""{base_injection}

# 馃摑 浠诲姟锛氫腑杞嫳缈昏瘧锛圕N鈫扙N ONLY锛?
## 杈撳叆鏂囨湰
{input_text}

## 杈撳嚭瑕佹眰
1. 鉁?浠呮墽琛?CN鈫扙N 璇█杞崲
2. 鉁?淇濇寔鍘熸剰瀹屾暣
3. 鉁?鏈鍑嗙‘鎬?
4. 鉁?瀛︽湳璇綋
5. 鉂?绂佹棰濆娑﹁壊
6. 鉂?绂佹鏀瑰彉鍘熸剰

## 杈撳嚭鏍煎紡
```
[缈昏瘧鍚庣殑鑻辨枃鏂囨湰]
```

## 鍏抽敭鏈瀵圭収
[鍒楀嚭5-10涓牳蹇冩湳璇殑缈昏瘧]
""",

        "鉁?琛ㄨ揪娑﹁壊": f"""{base_injection}

# 鉁?浠诲姟锛氳〃杈炬鼎鑹诧紙鍚岃瑷€浼樺寲锛?
## 杈撳叆鏂囨湰
{input_text}

## 娑﹁壊瑕佹眰
1. 鉁?鎻愬崌瀛︽湳鍦伴亾鎬?
2. 鉁?鏇挎崲楂樼骇鍔ㄨ瘝
3. 鉁?{"淇濇寔涓枃锛屼紭鍖栦腑鏂? if lang == "zh" else "淇濇寔鑻辨枃锛屼紭鍖栬嫳鏂?}
4. 鉁?淇濇寔鍘熸剰
5. 鉂?涓ユ牸绂佹缈昏瘧

## 杈撳嚭鏍煎紡
```
[娑﹁壊鍚庣殑鏂囨湰]
```

## 淇敼璇存槑
[鍒楀嚭涓昏淇敼鐐癸細鍔ㄨ瘝鏇挎崲銆佺粨鏋勮皟鏁寸瓑]
""",

        "馃攳 閫昏緫妫€鏌?: f"""{base_injection}

# 馃攳 浠诲姟锛氶€昏緫妫€鏌ワ紙涓嶄慨鏀规枃鏈級
## 杈撳叆鏂囨湰
{input_text}

## 妫€鏌ヨ鐐?
1. 鉁?鍥犳灉閾炬潯瀹屾暣鎬?
2. 鉁?璁烘嵁涓庤鐐逛竴鑷存€?
3. 鉁?閫昏緫琛旀帴娴佺晠鎬?
4. 鉁?鏈浣跨敤鍑嗙‘鎬?

## 杈撳嚭鏍煎紡
```
鉁?閫昏緫妫€鏌ラ€氳繃
```
鎴?
```
鉂?鍙戠幇浠ヤ笅闂:
1. [鍏蜂綋闂]
2. [鍏蜂綋闂]
...
```

**鈿狅笍 娉ㄦ剰**: 浠呮鏌ワ紝涓嶄慨鏀规枃鏈€?
""",

        "馃 鍘籄I鍛?(Humanizer)": f"""{base_injection}

# 馃 浠诲姟锛氬幓AI鍖栧鐞嗭紙闆剁炕璇戯級
## 杈撳叆鏂囨湰
{input_text}

## 鍘籄I鍖栬姹?
{"1. 鉁?浼樺寲涓枃璇簭锛屼繚鎸佷腑鏂囪緭鍑? if lang == "zh" else "1. 鉁?澧炲己鑻辨枃鑺傚鎰燂紝淇濇寔鑻辨枃杈撳嚭"}
2. 鉁?娑堥櫎 AI 甯哥敤璇?
3. 鉁?鍒犻櫎鏈烘杩炴帴璇?
4. 鉁?鑺傚鍙樺寲
5. 鉁?鎵胯涓嶇‘瀹氭€?
6. 鉂?涓ユ牸绂佹缈昏瘧

## 杈撳嚭鏍煎紡
```
[鍘籄I鍖栧悗鐨勬枃鏈琞
```

## 淇敼璇存槑
[鍒楀嚭3-5涓叧閿慨鏀圭偣鍙婄悊鐢盷
""",

        "馃懆鈥嶁殩锔?Reviewer瑙嗚": f"""{base_injection}

# 馃懆鈥嶁殩锔?浠诲姟锛歊eviewer瑙嗚瀹¤
## 杈撳叆鏂囨湰
{input_text}

## 瀹¤瑕佺偣
1. 鉁?瀛︽湳瑙勮寖鎬?
2. 鉁?璁鸿瘉寮哄害
3. 鉁?琛ㄨ揪娓呮櫚搴?
4. 鉁?娼滃湪闂

## 杈撳嚭鏍煎紡
## 馃摑 瀹＄鎰忚
[鎬讳綋璇勪环]

## 馃攳 鍏蜂綋寤鸿
1. [鍏蜂綋寤鸿]
2. [鍏蜂綋寤鸿]
...

## 鉁?浼樼偣
[鍒楀嚭浼樼偣]
""",

        "鉁嶏笍 閫愭璧疯崏": f"""{base_injection}

# 鉁嶏笍 浠诲姟锛氶€愭璧疯崏锛堝奖瀛愬啓浣滄ā寮忥級
## 杈撳叆澶х翰
{input_text}

## 璧疯崏瑕佹眰
1. 鉁?灏嗗ぇ绾叉墿鍏呬负姝ｅ紡娈佃惤
2. 鉁?妯′豢鏍囨潌鏂囩尞椋庢牸锛堝彞闀裤€佽繛鎺ヨ瘝銆佽鎬侊級
3. 鉁?绗﹀悎 {section} 鏉垮潡鐗圭偣
4. 鉁?淇濇寔瀛︽湳涓ヨ皑鎬?
5. 鉁?{"涓枃杈撳叆杈撳嚭涓枃锛岃嫳鏂囪緭鍏ヨ緭鍑鸿嫳鏂? if lang != "unknown" else "淇濇寔鍘熻瑷€"}

## 杈撳嚭鏍煎紡
```
[鎵╁厖鍚庣殑姝ｅ紡娈佃惤]
```

## 妯′豢璇存槑
[璇存槑妯′豢浜嗘爣鏉嗘枃鐚殑鍝簺鐗瑰緛锛氬彞闀裤€佽繛鎺ヨ瘝銆佽鎬佺瓑]
""",

        "馃幆 绮句慨妯″紡": f"""{base_injection}

# 馃幆 浠诲姟锛氭繁搴︾簿淇紙椤跺垔姘村噯锛?
## 杈撳叆鏂囨湰
{input_text}

## 绮句慨瑕佹眰
1. 鉁?杈惧埌椤跺垔鍑虹増姘村噯
2. 鉁?妯′豢鏍囨潌鏂囩尞椋庢牸
3. 鉁?绗﹀悎 {section} 鏉垮潡鐗圭偣
4. 鉁?鏈绮剧‘浣跨敤
5. 鉁?{"淇濇寔鍘熻瑷€锛屼紭鍖栧師璇█" if lang != "unknown" else "淇濇寔鍘熻瑷€"}

## 杈撳嚭鏍煎紡
```
[绮句慨鍚庣殑鏂囨湰]
```

## 淇敼璇存槑
[璇︾粏璇存槑淇敼鍐呭鍜岀悊鐢盷
""",

        "馃摑 寮曠敤楠岃瘉": f"""{base_injection}

# 馃摑 浠诲姟锛氬紩鐢ㄦ牸寮忛獙璇?
## 杈撳叆鏂囨湰
{input_text}

## 楠岃瘉瑕佺偣
1. 鉁?寮曠敤鏍煎紡缁熶竴
2. 鉁?寮曠敤瀹屾暣鎬?
3. 鉁?寮曠敤鐩稿叧鎬?

## 杈撳嚭鏍煎紡
```
鉁?寮曠敤鏍煎紡楠岃瘉閫氳繃
```
鎴?
```
鉂?鍙戠幇浠ヤ笅闂:
1. [鍏蜂綋闂]
2. [鍏蜂綋闂]
...
```
""",

        "馃帹 鍥捐〃瑙勮寖妫€鏌?: f"""{base_injection}

# 馃帹 浠诲姟锛氬浘琛ㄦ弿杩拌鑼冩鏌?
## 杈撳叆鏂囨湰
{input_text}

## 妫€鏌ヨ鐐?
1. 鉁?鍥捐〃寮曠敤瀹屾暣鎬?
2. 鉁?鍥捐〃鎻忚堪娓呮櫚搴?
3. 鉁?鑹茬洸鍙嬪ソ鎬?
4. 鉁?鍧愭爣杞存爣娉?

## 杈撳嚭鏍煎紡
```
鉁?鍥捐〃鎻忚堪瑙勮寖
```
鎴?
```
鉂?鍙戠幇浠ヤ笅闂:
1. [鍏蜂綋闂]
2. [鍏蜂綋闂]
...
```
""",

        "馃搵 Redlining淇": f"""{base_injection}

# 馃搵 浠诲姟锛歊edlining淇锛堟樉绀轰慨鏀圭棔杩癸級
## 杈撳叆鏂囨湰
{input_text}

## 淇瑕佹眰
1. 鉁?鏄剧ず淇敼鐥曡抗
2. 鉁?鎻愪緵瀵规瘮瑙嗗浘
3. 鉁?淇濇寔鍘熻瑷€
4. 鉁?璇存槑淇敼鐞嗙敱

## 杈撳嚭鏍煎紡
```
## 淇敼鍚庢枃鏈?
[淇敼鍚庣殑鏂囨湰]

## 淇敼璇存槑
[璇︾粏鍒楀嚭鎵€鏈変慨鏀圭偣]
```
""",
    }

    # 榛樿妯℃澘
    default_template = f"""{base_injection}

# 馃幆 浠诲姟锛歿function}
## 杈撳叆鏂囨湰
{input_text}

## 杈撳嚭瑕佹眰
涓ユ牸鎸夌収 {function} 鍔熻兘瑙勫垯鎵ц
{"淇濇寔鍘熻瑷€锛岀粷涓嶇炕璇? if lang != "unknown" else ""}

## 杈撳嚭鏍煎紡
```
[澶勭悊缁撴灉]
```

## 淇敼璇存槑
[绠€瑕佽鏄庡鐞嗗唴瀹筣
"""

    return function_templates.get(function, default_template)


# 鈹€鈹€ 鍘熺敓 API 璋冪敤锛堝甫閲嶈瘯锛夆攢鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

def get_client():
    if not CLAUDE_API_KEY:
        st.error("鉂?鏈娴嬪埌 ANTHROPIC_AUTH_TOKEN")
        st.stop()
    return anthropic.Anthropic(
        api_key=CLAUDE_API_KEY,
        base_url=CLAUDE_BASE_URL if CLAUDE_BASE_URL != "https://api.anthropic.com" else None
    )


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    retry=retry_if_exception_type((anthropic.APITimeoutError, anthropic.InternalServerError)),
)
def call_api(prompt: str, timeout: int = 180) -> str:
    client = get_client()
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        temperature=0.2,
        timeout=timeout,
        messages=[{"role": "user", "content": prompt}]
    )
    return message.content[0].text


# 鈹€鈹€ 鏂囦欢澶勭悊 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

def extract_text(file_bytes: bytes, filename: str) -> str:
    """鎻愬彇鏂囦欢鏂囨湰"""
    try:
        if filename.endswith('.pdf'):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            return "\n\n".join([p.get_text() for p in doc])
        elif filename.endswith('.docx'):
            doc = Document(BytesIO(file_bytes))
            return "\n\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        return f"瑙ｆ瀽澶辫触: {e}"
    return ""


def create_docx_with_redlines(original: str, revised: str, metadata: dict) -> bytes:
    """鍒涘缓甯︿慨璁㈢棔杩圭殑Word鏂囨。"""
    doc = Document()
    doc.add_heading(f"瀛︾爺路淇妯″紡 - {metadata['function']}", 0)

    # 鍏冧俊鎭?
    p = doc.add_paragraph()
    p.add_run(f"鏉垮潡: {metadata['section']}\n")
    p.add_run(f"棰嗗煙: {metadata['domain']}\n")
    p.add_run(f"鏃堕棿: {metadata['timestamp']}\n")

    doc.add_heading("鍘熸枃", 1)
    doc.add_paragraph(original)

    doc.add_heading("淇敼鍚?, 1)
    doc.add_paragraph(revised)

    # 绠€鍗曞樊寮傛爣娉紙浣跨敤棰滆壊锛?
    doc.add_heading("淇敼璇存槑", 2)
    original_words = set(original.lower().split())
    revised_words = set(revised.lower().split())

    added = revised_words - original_words
    removed = original_words - revised_words

    if added:
        p = doc.add_paragraph()
        run = p.add_run("鏂板璇嶆眹: ")
        run.font.color.rgb = RGBColor(0, 128, 0)
        p.add_run(", ".join(list(added)[:20]))

    if removed:
        p = doc.add_paragraph()
        run = p.add_run("鍒犻櫎璇嶆眹: ")
        run.font.color.rgb = RGBColor(255, 0, 0)
        p.add_run(", ".join(list(removed)[:20]))

    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.read()


def create_docx(content: str, metadata: dict) -> bytes:
    """鍒涘缓 Word 鏂囨。"""
    doc = Document()
    doc.add_heading(f"瀛︾爺路宸ョ绉戠爺鍔╂墜 - {metadata['function']}", 0)

    # 鍏冧俊鎭?
    p = doc.add_paragraph()
    p.add_run(f"鏉垮潡: {metadata['section']}\n")
    p.add_run(f"棰嗗煙: {metadata['domain']}\n")
    p.add_run(f"鏃堕棿: {metadata['timestamp']}\n")

    doc.add_heading("澶勭悊缁撴灉", 1)
    doc.add_paragraph(content)

    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.read()


# 鈹€鈹€ UI Styling 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%); }
    .main-header {
        text-align: center; padding: 2rem 0;
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white; border-radius: 12px; margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.2);
    }
    .section-tab { background: white; padding: 0.5rem 1rem; border-radius: 8px; font-weight: 600; }
    .comparison-box { background: white; padding: 1.5rem; border-radius: 8px; border-left: 4px solid #3b82f6; min-height: 400px; }
    .stButton>button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb, #3b82f6);
        color: white; border: none; padding: 0.75rem 2rem;
        font-weight: 600; border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)


# 鈹€鈹€ Header 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

st.markdown("""
<div class="main-header">
    <h1>馃И 瀛︾爺路宸ョ绉戠爺鍔╂墜 v4.0</h1>
    <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">15+鍔熻兘鐭╅樀 路 鏉垮潡閿氬畾 路 闆剁浣嶆墽琛?路 褰卞瓙鍚堣憲鑰?路 璇█鍩哄洜娣卞害鎻愬彇</p>
</div>
""", unsafe_allow_html=True)

# API 鐘舵€佹爮
col1, col2, col3 = st.columns(3)
with col1:
    st.success(f"鉁?API 灏辩华" if CLAUDE_API_KEY else "鉂?API 鏈厤缃?)
with col2:
    st.info(f"馃摎 {len(LOCAL_SKILLS)} 涓鏈?Skills")
with col3:
    st.success("鈴憋笍 180s 瓒呮椂 | 3娆￠噸璇?)

# 鈹€鈹€ 渚ц竟鏍忥細15+鍔熻兘鐭╅樀 + 褰卞瓙鍚堣憲鑰?鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

with st.sidebar:
    st.markdown("---")
    st.subheader("馃幆 15+ 鍔熻兘鐭╅樀")

    function = st.selectbox(
        "閫夋嫨鍔熻兘",
        list(FUNCTION_MATRIX.keys()),
        help="姣忎釜鍔熻兘涓ユ牸鍗曚竴鑱岃矗锛岄浂绡′綅鎵ц"
    )

    # 鏄剧ず褰撳墠鍔熻兘瑙勫垯
    func_info = FUNCTION_MATRIX.get(function, {})
    if func_info:
        st.caption(f"馃搵 {func_info.get('description', '')}")
        for rule in func_info.get('rules', []):
            st.caption(rule)

    st.markdown("---")
    st.subheader("馃敩 瀛︾棰嗗煙")

    domain = st.selectbox(
        "鐮旂┒棰嗗煙",
        list(DOMAINS.keys()),
        help="閫夋嫨棰嗗煙浠ユ縺娲绘湳璇‖閿佷繚鎶?
    )

    st.markdown("---")
    st.subheader("馃摎 褰卞瓙鍚堣憲鑰咃細鏍囨潌鏂囩尞")

    reference_files = st.file_uploader(
        "涓婁紶鏍囨潌鏂囩尞 (1-3绡?",
        type=["pdf", "docx"],
        accept_multiple_files=True,
        help="AI灏嗗涔犲叾鍐欎綔椋庢牸骞舵ā浠?
    )

    reference_styles = []
    if reference_files:
        for ref_file in reference_files:
            if ref_file.name not in st.session_state.reference_docs:
                with st.spinner(f"鍒嗘瀽 {ref_file.name}..."):
                    file_bytes = ref_file.read()
                    text = extract_text(file_bytes, ref_file.name)
                    if text and not text.startswith("瑙ｆ瀽"):
                        style = analyze_reference_paper(file_bytes, ref_file.name)
                        st.session_state.reference_docs[ref_file.name] = style
                        st.success(f"鉁?{ref_file.name}")

        # 鏀堕泦鎵€鏈夋爣鏉嗘枃鐚鏍?
        reference_styles = list(st.session_state.reference_docs.values())

    if st.session_state.reference_docs:
        st.success(f"鉁?宸插姞杞?{len(st.session_state.reference_docs)} 绡囨爣鏉嗘枃鐚?)

    st.markdown("---")
    st.subheader("馃搧 鏅€氭枃浠朵笂浼?)

    uploaded_file = st.file_uploader("涓婁紶寰呭鐞嗘枃浠?, type=["pdf", "docx"])

    if uploaded_file:
        extracted = extract_text(uploaded_file.read(), uploaded_file.name)
        if extracted and not extracted.startswith("瑙ｆ瀽"):
            st.success(f"鉁?鎻愬彇 {len(extracted)} 瀛楃")
            if st.button("馃摜 濉叆缂栬緫鍣?):
                st.session_state.current_input = extracted[:10000]
                st.rerun()

# 鈹€鈹€ 椤堕儴锛氭澘鍧楅敋鐐?Tabs 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

st.markdown("---")
section_tabs = st.tabs([
    "馃摑 鎽樿",
    "馃摌 寮曡█",
    "馃敩 鏂规硶",
    "馃搳 缁撴灉",
    "馃挕 璁ㄨ",
    "馃弫 缁撹"
])

# 纭畾褰撳墠鏉垮潡
current_section = section_names[0]
for i, tab in enumerate(section_tabs):
    with tab:
        st.session_state[f"active_section_{i}"] = section_names[i]

# 閫氳繃 query params 鎴?session state 纭畾褰撳墠婵€娲荤殑 tab
# Streamlit tabs 鏃犳硶鐩存帴鑾峰彇婵€娲荤储寮曪紝鐢?radio 闅愬紡璺熻釜
if "selected_section_idx" not in st.session_state:
    st.session_state.selected_section_idx = 0

for i, tab in enumerate(section_tabs):
    with tab:
        if st.button("馃搶 閿氬畾姝ゆ澘鍧?, key=f"anchor_{i}", help="鐐瑰嚮鍚庝笅鏂规搷浣滃皢閽堝姝ゆ澘鍧?):
            st.session_state.selected_section_idx = i
            st.rerun()

current_section = section_names[st.session_state.selected_section_idx]

# 鈹€鈹€ 涓诲尯鍩燂細宸﹀彸鍒嗘爮瀵规瘮 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

st.markdown("---")
col_left, col_right = st.columns([1, 1])

with col_left:
    st.markdown(f"### 馃摑 鍘熸枃杈撳叆 [{current_section}]")
    current_input = st.text_area(
        "",
        value=st.session_state.get("current_input", ""),
        height=400,
        label_visibility="collapsed",
        key=f"input_{current_section}"
    )
    st.session_state.current_input = current_input

    st.caption(f"馃搳 {len(current_input)} 瀛楃 | 璇█: {'涓枃' if detect_language(current_input) == 'zh' else '鑻辨枃'}")

    st.markdown("---")
    btn_col1, btn_col2, btn_col3 = st.columns([3, 1, 1])
    with btn_col1:
        process_button = st.button(
            f"鉁?鎵ц {function}",
            type="primary",
            use_container_width=True
        )
    with btn_col2:
        if st.button("馃棏锔?):
            st.session_state.current_input = ""
            st.rerun()
    with btn_col3:
        if st.button("馃攧"):
            st.rerun()

with col_right:
    st.markdown(f"### 馃憗锔?澶勭悊缁撴灉 [{current_section}]")
    result_placeholder = st.empty()
    modification_note = st.empty()

    if process_button and current_input.strip():
        input_lang = detect_language(current_input)

        with st.spinner(f"馃攧 澶勭悊涓?[{function}] | 鏉垮潡: {current_section} | 棰嗗煙: {domain} | 璇█: {'涓枃' if input_lang == 'zh' else '鑻辨枃'}..."):
            try:
                # 纭攣淇濇姢
                protected_input, term_mapping = protect_hard_terms(current_input, domain)

                # 鍒涘缓闆剁浣?prompt
                full_prompt = create_strict_prompt(
                    function, protected_input, current_section, domain, reference_styles, input_lang
                )

                # 璋冪敤 API
                start_time = time.time()
                response = call_api(full_prompt, timeout=180)
                elapsed = time.time() - start_time

                if response:
                    # 鎭㈠鏈
                    final_output = restore_hard_terms(response, term_mapping)

                    # 淇濆瓨鍒板巻鍙?
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    history_entry = {
                        "id": len(st.session_state.history) + 1,
                        "timestamp": timestamp,
                        "function": function,
                        "section": current_section,
                        "domain": domain,
                        "input_lang": input_lang,
                        "input": current_input,
                        "output": final_output,
                        "elapsed": f"{elapsed:.1f}s"
                    }
                    st.session_state.history.insert(0, history_entry)
                    st.session_state.history = st.session_state.history[:100]

                    # 鏄剧ず缁撴灉
                    result_placeholder.markdown(final_output)

                    # 淇敼璇存槑
                    modification_note.caption(
                        f"鉁?宸查拡瀵广€恵current_section}銆戝畬鎴愩€恵function}銆戯紝"
                        f"涓昏浼樺寲浜嗚搴忓拰琛ㄨ揪锛屼娇璁鸿瘉鏇村叿{'鐪熶汉鑺傚鎰? if 'Humanizer' in function else '瀛︽湳涓撲笟鎬?}"
                    )

                    # 瀵煎嚭鎸夐挳锛堟敮鎸乺edlining锛?
                    st.markdown("---")
                    export_col1, export_col2 = st.columns(2)

                    # 鏍囧噯瀵煎嚭
                    doc_bytes = create_docx(final_output, {
                        "function": function,
                        "section": current_section,
                        "domain": domain,
                        "timestamp": timestamp
                    })
                    with export_col1:
                        st.download_button(
                            "馃摜 瀵煎嚭 Word",
                            doc_bytes,
                            file_name=f"yanyu_{current_section}_{timestamp.replace(':', '-')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

                    # Redlining瀵煎嚭锛堜粎闄愪慨鏀圭被鍔熻兘锛?
                    if function in ["鉁?琛ㄨ揪娑﹁壊", "馃 鍘籄I鍛?(Humanizer)", "馃幆 绮句慨妯″紡"]:
                        redline_bytes = create_docx_with_redlines(current_input, final_output, {
                            "function": function,
                            "section": current_section,
                            "domain": domain,
                            "timestamp": timestamp
                        })
                        with export_col2:
                            st.download_button(
                                "馃搵 瀵煎嚭 Redlining",
                                redline_bytes,
                                file_name=f"yanyu_redline_{current_section}_{timestamp.replace(':', '-')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

            except Exception as e:
                st.error(f"鉂?澶勭悊澶辫触: {e}")
                if "timeout" in str(e).lower():
                    st.warning("馃挕 寤鸿锛氱缉鐭枃鏈垨澧炲姞瓒呮椂鏃堕棿")

# 鈹€鈹€ 鐗堟湰鏃跺厜鏈?鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

st.markdown("---")
st.subheader("鈴?鐗堟湰鏃跺厜鏈?)

# 鎸夋澘鍧楃瓫閫?
filter_section = st.selectbox(
    "绛涢€夋澘鍧?,
    ["鍏ㄩ儴"] + list(SECTIONS.keys()),
    index=0
)

# 鎸夊姛鑳界瓫閫?
filter_function = st.selectbox(
    "绛涢€夊姛鑳?,
    ["鍏ㄩ儴"] + list(FUNCTION_MATRIX.keys()),
    index=0
)

# 杩囨护鍘嗗彶
filtered_history = st.session_state.history
if filter_section != "鍏ㄩ儴":
    filtered_history = [h for h in filtered_history if h["section"] == filter_section]
if filter_function != "鍏ㄩ儴":
    filtered_history = [h for h in filtered_history if h["function"] == filter_function]

if filtered_history:
    for entry in filtered_history[:20]:
        with st.expander(
            f"#{entry['id']} 路 {entry['function']} 路 {entry['section']} 路 {entry['timestamp']}",
            expanded=False
        ):
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.caption(f"棰嗗煙: {entry['domain']} | 璇█: {'涓枃' if entry['input_lang'] == 'zh' else '鑻辨枃'} | 鑰楁椂: {entry['elapsed']}")
            with col2:
                if st.button("馃摜 鎭㈠", key=f"restore_{entry['id']}"):
                    st.session_state.current_input = entry['input']
                    st.rerun()
            with col3:
                if st.button("馃棏锔?, key=f"delete_{entry['id']}"):
                    st.session_state.history = [h for h in st.session_state.history if h['id'] != entry['id']]
                    st.rerun()

            # 瀵规瘮瑙嗗浘
            col_orig, col_out = st.columns(2)
            with col_orig:
                st.markdown("**鍘熷杈撳叆**")
                st.text_area("", entry['input'], height=150, key=f"orig_{entry['id']}", disabled=True)
            with col_out:
                st.markdown("**澶勭悊杈撳嚭**")
                st.markdown(entry['output'])
            st.markdown("---")
else:
    st.info("馃摥 鏆傛棤绗﹀悎鏉′欢鐨勮褰?)

# 鈹€鈹€ Footer 鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€鈹€

st.markdown("""
---
<div style="text-align: center; color: #64748b; font-size: 0.8rem; padding: 2rem 0;">
    <p><strong>馃И 瀛︾爺路宸ョ绉戠爺鍔╂墜 v4.0</strong></p>
    <p>15+鍔熻兘鐭╅樀 路 鏉垮潡閿氬畾 路 闆剁浣嶆墽琛?路 褰卞瓙鍚堣憲鑰?路 璇█鍩哄洜娣卞害鎻愬彇</p>
    <p>馃敩 9澶у绉戦鍩?路 馃摎 鏈湴瀛︽湳搴撻泦鎴?路 鈴?鐗堟湰鏃跺厜鏈?路 馃搵 Redlining鏀寔</p>
</div>
""", unsafe_allow_html=True)
