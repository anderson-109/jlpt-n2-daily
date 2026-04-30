"""
JLPT N2 Daily Word Document Generator
每天自動從題庫隨機挑選 3 組單字 + 3 組語法，生成帶振假名的 Word 檔
"""

import random
import zipfile
import shutil
import os
import re
from datetime import datetime, timezone, timedelta

# ── 題庫 ──────────────────────────────────────────────────────────────────────

VOCAB_BANK = [
    {
        "word": "遠慮", "reading": "えんりょ", "pos": "名・する",
        "meaning": "客氣；謝絕；顧慮",
        "sentence": "遠慮なく召し上がってください。",
        "translation": "請不要客氣，盡情享用。",
        "rubies": [("遠慮", "えんりょ"), ("召", "め"), ("上", "あ")],
        "plain_parts": ["なく", "し", "がってください。"],
    },
    {
        "word": "促す", "reading": "うながす", "pos": "動（他）",
        "meaning": "催促；促進；提醒",
        "sentence": "先生は生徒に発言を促した。",
        "translation": "老師促使學生發言。",
        "rubies": [("先生", "せんせい"), ("生徒", "せいと"), ("発言", "はつげん"), ("促", "うなが")],
        "plain_parts": ["は", "に", "を", "した。"],
    },
    {
        "word": "承る", "reading": "うけたまわる", "pos": "動（他）",
        "meaning": "（謙）接受；聆聽；承接",
        "sentence": "ご注文を承ります。",
        "translation": "我來為您接受點餐（敬語用法）。",
        "rubies": [("注文", "ちゅうもん"), ("承", "うけたまわ")],
        "plain_parts": ["ご", "を", "ります。"],
        "prefix": "ご",
    },
    {
        "word": "把握", "reading": "はあく", "pos": "名・する",
        "meaning": "掌握；理解；抓住",
        "sentence": "状況を把握してから行動してください。",
        "translation": "請先掌握狀況後再行動。",
        "rubies": [("状況", "じょうきょう"), ("把握", "はあく"), ("行動", "こうどう")],
        "plain_parts": ["を", "してから", "してください。"],
    },
    {
        "word": "緊張", "reading": "きんちょう", "pos": "名・する",
        "meaning": "緊張；拉緊",
        "sentence": "発表の前に緊張してしまった。",
        "translation": "在發表前不禁緊張起來了。",
        "rubies": [("発表", "はっぴょう"), ("前", "まえ"), ("緊張", "きんちょう")],
        "plain_parts": ["の", "に", "してしまった。"],
    },
    {
        "word": "影響", "reading": "えいきょう", "pos": "名・する",
        "meaning": "影響",
        "sentence": "台風の影響で電車が止まった。",
        "translation": "受颱風影響，電車停駛了。",
        "rubies": [("台風", "たいふう"), ("影響", "えいきょう"), ("電車", "でんしゃ")],
        "plain_parts": ["の", "で", "が止まった。"],
    },
    {
        "word": "我慢", "reading": "がまん", "pos": "名・する",
        "meaning": "忍耐；克制",
        "sentence": "痛くても我慢しなければならない。",
        "translation": "即使很痛也不得不忍耐。",
        "rubies": [("痛", "いた"), ("我慢", "がまん")],
        "plain_parts": ["くても", "しなければならない。"],
    },
    {
        "word": "感謝", "reading": "かんしゃ", "pos": "名・する",
        "meaning": "感謝；感恩",
        "sentence": "助けてくれた人に感謝する。",
        "translation": "向幫助自己的人表示感謝。",
        "rubies": [("助", "たす"), ("人", "ひと"), ("感謝", "かんしゃ")],
        "plain_parts": ["けてくれた", "に", "する。"],
    },
    {
        "word": "判断", "reading": "はんだん", "pos": "名・する",
        "meaning": "判斷；決定",
        "sentence": "自分で判断することが大切だ。",
        "translation": "自己做判斷是很重要的。",
        "rubies": [("自分", "じぶん"), ("判断", "はんだん"), ("大切", "たいせつ")],
        "plain_parts": ["で", "することが", "だ。"],
    },
    {
        "word": "丁寧", "reading": "ていねい", "pos": "な形",
        "meaning": "仔細；有禮貌；認真",
        "sentence": "丁寧に説明してくれてありがとう。",
        "translation": "謝謝你仔細地說明。",
        "rubies": [("丁寧", "ていねい"), ("説明", "せつめい")],
        "plain_parts": ["に", "してくれてありがとう。"],
    },
    {
        "word": "確認", "reading": "かくにん", "pos": "名・する",
        "meaning": "確認；核實",
        "sentence": "メールを送る前に内容を確認する。",
        "translation": "在發送郵件前確認內容。",
        "rubies": [("前", "まえ"), ("内容", "ないよう"), ("確認", "かくにん")],
        "plain_parts": ["メールを送る", "に", "を", "する。"],
    },
    {
        "word": "挑戦", "reading": "ちょうせん", "pos": "名・する",
        "meaning": "挑戰",
        "sentence": "新しいことに挑戦するのは楽しい。",
        "translation": "挑戰新事物是件快樂的事。",
        "rubies": [("挑戦", "ちょうせん"), ("楽", "たの")],
        "plain_parts": ["新しいことに", "するのは", "しい。"],
    },
]

GRAMMAR_BANK = [
    {
        "pattern": "〜にもかかわらず",
        "explanation": "儘管…；雖然…卻…\n表示逆接，前後結果與預期相反",
        "connection": "名詞／動詞普通形＋にもかかわらず",
        "sentence": "雨にもかかわらず、試合は行われた。",
        "translation": "儘管下著雨，比賽還是照常舉行了。",
        "rubies": [("雨", "あめ"), ("試合", "しあい"), ("行", "おこな")],
        "plain_parts": ["にもかかわらず、", "は", "われた。"],
    },
    {
        "pattern": "〜わけにはいかない",
        "explanation": "不能…；不可以…\n表示基於道義、常識或規定，無法做某事",
        "connection": "動詞辭書形／ない形＋わけにはいかない",
        "sentence": "大事な会議があるので、休むわけにはいかない。",
        "translation": "因為有重要的會議，所以不能請假。",
        "rubies": [("大事", "だいじ"), ("会議", "かいぎ"), ("休", "やす")],
        "plain_parts": ["な", "があるので、", "むわけにはいかない。"],
    },
    {
        "pattern": "〜に反して",
        "explanation": "與…相反；違反…\n表示結果或行為與某預期、規定相反",
        "connection": "名詞＋に反して",
        "sentence": "予想に反して、試験はとても簡単だった。",
        "translation": "與預期相反，考試非常簡單。",
        "rubies": [("予想", "よそう"), ("反", "はん"), ("試験", "しけん"), ("簡単", "かんたん")],
        "plain_parts": ["に", "して、", "はとても", "だった。"],
    },
    {
        "pattern": "〜ざるを得ない",
        "explanation": "不得不…；只好…\n表示在某種情況下別無選擇",
        "connection": "動詞ない形（ない→ず）＋ざるを得ない",
        "sentence": "体調が悪くて、会社を休まざるを得なかった。",
        "translation": "身體不舒服，不得不向公司請假。",
        "rubies": [("体調", "たいちょう"), ("悪", "わる"), ("会社", "かいしゃ"), ("休", "やす")],
        "plain_parts": ["が", "くて、", "を", "まざるを得なかった。"],
    },
    {
        "pattern": "〜に過ぎない",
        "explanation": "只不過是…；僅僅是…\n強調程度輕微或範圍有限",
        "connection": "名詞／動詞普通形＋に過ぎない",
        "sentence": "これは私の意見に過ぎない。",
        "translation": "這只不過是我個人的意見而已。",
        "rubies": [("私", "わたし"), ("意見", "いけん"), ("過", "す")],
        "plain_parts": ["これは", "の", "に", "ぎない。"],
    },
    {
        "pattern": "〜をきっかけに",
        "explanation": "以…為契機；以…為起點\n表示某事成為轉變或開始的機會",
        "connection": "名詞＋をきっかけに",
        "sentence": "留学をきっかけに、日本語が好きになった。",
        "translation": "以留學為契機，開始喜歡上日語。",
        "rubies": [("留学", "りゅうがく"), ("日本語", "にほんご"), ("好", "す")],
        "plain_parts": ["をきっかけに、", "が", "きになった。"],
    },
    {
        "pattern": "〜かねない",
        "explanation": "可能…；說不定會…\n表示有發生某種不好結果的可能性",
        "connection": "動詞ます形＋かねない",
        "sentence": "無理をすると、体を壊しかねない。",
        "translation": "如果勉強的話，可能會搞壞身體。",
        "rubies": [("無理", "むり"), ("体", "からだ"), ("壊", "こわ")],
        "plain_parts": ["をすると、", "を", "しかねない。"],
    },
    {
        "pattern": "〜に加えて",
        "explanation": "除…之外還…；再加上…\n表示在某事物基礎上又添加了新的內容",
        "connection": "名詞＋に加えて",
        "sentence": "実力に加えて、運も必要だ。",
        "translation": "除了實力之外，運氣也是必要的。",
        "rubies": [("実力", "じつりょく"), ("加", "くわ"), ("運", "うん"), ("必要", "ひつよう")],
        "plain_parts": ["に", "えて、", "も", "だ。"],
    },
    {
        "pattern": "〜からといって",
        "explanation": "雖說…但…；即使…也不能…\n表示不能僅憑某理由就做某事",
        "connection": "動詞／形容詞普通形＋からといって",
        "sentence": "忙しいからといって、食事を抜いてはいけない。",
        "translation": "就算再忙，也不能不吃飯。",
        "rubies": [("忙", "いそが"), ("食事", "しょくじ"), ("抜", "ぬ")],
        "plain_parts": ["しいからといって、", "を", "いてはいけない。"],
    },
    {
        "pattern": "〜ものの",
        "explanation": "雖然…但是…\n表示承認前項事實，但後項與預期不符",
        "connection": "動詞／形容詞普通形＋ものの",
        "sentence": "日本語を勉強したものの、まだ話せない。",
        "translation": "雖然學了日語，但還是不會說。",
        "rubies": [("日本語", "にほんご"), ("勉強", "べんきょう"), ("話", "はな")],
        "plain_parts": ["を", "したものの、まだ", "せない。"],
    },
]

# ── XML helpers ───────────────────────────────────────────────────────────────

def rpr(bold=False, size=18):
    b = "<w:b/><w:bCs/>" if bold else "<w:b w:val=\"false\"/><w:bCs w:val=\"false\"/>"
    return (f"<w:rPr>"
            f"<w:rFonts w:ascii=\"Arial\" w:cs=\"Arial\" w:eastAsia=\"Arial\" w:hAnsi=\"Arial\"/>"
            f"{b}<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/></w:rPr>")

def rt_rpr():
    return (f"<w:rPr>"
            f"<w:rFonts w:ascii=\"Arial\" w:cs=\"Arial\" w:eastAsia=\"Arial\" w:hAnsi=\"Arial\"/>"
            f"<w:sz w:val=\"12\"/><w:szCs w:val=\"12\"/></w:rPr>")

def ruby(base, reading, bold=False):
    return (f"<w:ruby>"
            f"<w:rubyPr>"
            f"<w:rubyAlign w:val=\"distributeSpace\"/>"
            f"<w:hps w:val=\"12\"/><w:hpsRaise w:val=\"20\"/>"
            f"<w:hpsBaseText w:val=\"18\"/><w:lid w:val=\"ja-JP\"/>"
            f"</w:rubyPr>"
            f"<w:rt><w:r>{rt_rpr()}<w:t>{reading}</w:t></w:r></w:rt>"
            f"<w:rubyBase><w:r>{rpr(bold)}<w:t>{base}</w:t></w:r></w:rubyBase>"
            f"</w:ruby>")

def run(text, bold=False):
    sp = " xml:space=\"preserve\"" if (" " in text or text != text.strip()) else ""
    return f"<w:r>{rpr(bold)}<w:t{sp}>{text}</w:t></w:r>"

def build_sentence_xml(sentence, rubies, plain_parts):
    """Build interleaved ruby+plain runs for a sentence."""
    result = ""
    remaining = sentence
    for i, (kanji, reading) in enumerate(rubies):
        idx = remaining.find(kanji)
        if idx > 0:
            result += run(remaining[:idx])
        result += ruby(kanji, reading)
        remaining = remaining[idx + len(kanji):]
    if remaining:
        result += run(remaining)
    return result

# ── DOCX builder ─────────────────────────────────────────────────────────────

BORDER = '<w:top w:val="single" w:color="CCCCCC" w:sz="4"/><w:left w:val="single" w:color="CCCCCC" w:sz="4"/><w:bottom w:val="single" w:color="CCCCCC" w:sz="4"/><w:right w:val="single" w:color="CCCCCC" w:sz="4"/>'
MARGINS = '<w:top w:type="dxa" w:w="100"/><w:left w:type="dxa" w:w="120"/><w:bottom w:type="dxa" w:w="100"/><w:right w:type="dxa" w:w="120"/>'

def tc(content_xml, width, fill, valign="center"):
    return (f"<w:tc>"
            f"<w:tcPr>"
            f"<w:tcW w:type=\"dxa\" w:w=\"{width}\"/>"
            f"<w:tcBorders>{BORDER}</w:tcBorders>"
            f"<w:shd w:fill=\"{fill}\" w:val=\"clear\"/>"
            f"<w:tcMar>{MARGINS}</w:tcMar>"
            f"<w:vAlign w:val=\"{valign}\"/>"
            f"</w:tcPr>"
            f"<w:p>{content_xml}</w:p>"
            f"</w:tc>")

def header_tc(text, width, fill):
    content = (f"<w:pPr><w:jc w:val=\"center\"/></w:pPr>"
               f"<w:r><w:rPr>"
               f"<w:rFonts w:ascii=\"Arial\" w:cs=\"Arial\" w:eastAsia=\"Arial\" w:hAnsi=\"Arial\"/>"
               f"<w:b/><w:bCs/><w:color w:val=\"FFFFFF\"/>"
               f"<w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr>"
               f"<w:t>{text}</w:t></w:r>")
    return tc(content, width, fill)

def plain_tc(text, width, fill, bold=False, center=False):
    align = "<w:pPr><w:jc w:val=\"center\"/></w:pPr>" if center else ""
    content = align + run(text, bold)
    return tc(content, width, fill)

def ruby_tc(sentence_xml, width, fill):
    return tc(sentence_xml, width, fill)

def multiline_tc(text, width, fill):
    lines = text.split("\n")
    paragraphs = ""
    for line in lines:
        paragraphs += f"<w:p>{run(line)}</w:p>"
    return (f"<w:tc>"
            f"<w:tcPr>"
            f"<w:tcW w:type=\"dxa\" w:w=\"{width}\"/>"
            f"<w:tcBorders>{BORDER}</w:tcBorders>"
            f"<w:shd w:fill=\"{fill}\" w:val=\"clear\"/>"
            f"<w:tcMar>{MARGINS}</w:tcMar>"
            f"<w:vAlign w:val=\"center\"/>"
            f"</w:tcPr>"
            f"{paragraphs}"
            f"</w:tc>")

def heading_para(text, level=1, color="2E5FA3", size=36, align="left"):
    style = f"Heading{level}"
    jc = f"<w:jc w:val=\"{align}\"/>" if align != "left" else ""
    return (f"<w:p>"
            f"<w:pPr><w:pStyle w:val=\"{style}\"/>{jc}</w:pPr>"
            f"<w:r><w:rPr>"
            f"<w:rFonts w:ascii=\"Arial\" w:cs=\"Arial\" w:eastAsia=\"Arial\" w:hAnsi=\"Arial\"/>"
            f"<w:b/><w:bCs/><w:color w:val=\"{color}\"/>"
            f"<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/></w:rPr>"
            f"<w:t>{text}</w:t></w:r>"
            f"</w:p>")

def subtitle_para(text):
    return (f"<w:p>"
            f"<w:pPr><w:jc w:val=\"center\"/>"
            f"<w:spacing w:after=\"400\"/></w:pPr>"
            f"<w:r><w:rPr>"
            f"<w:rFonts w:ascii=\"Arial\" w:cs=\"Arial\" w:eastAsia=\"Arial\" w:hAnsi=\"Arial\"/>"
            f"<w:color w:val=\"888888\"/><w:i/><w:iCs/>"
            f"<w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr>"
            f"<w:t>{text}</w:t></w:r>"
            f"</w:p>")

def spacer():
    return "<w:p><w:pPr><w:spacing w:before=\"300\" w:after=\"100\"/></w:pPr></w:p>"

def footer_para(text):
    return (f"<w:p>"
            f"<w:pPr><w:jc w:val=\"center\"/>"
            f"<w:spacing w:before=\"400\"/></w:pPr>"
            f"<w:r><w:rPr>"
            f"<w:rFonts w:ascii=\"Arial\" w:cs=\"Arial\" w:eastAsia=\"Arial\" w:hAnsi=\"Arial\"/>"
            f"<w:color w:val=\"888888\"/><w:i/><w:iCs/>"
            f"<w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr>"
            f"<w:t>{text}</w:t></w:r>"
            f"</w:p>")

# ── Table builders ────────────────────────────────────────────────────────────

VOCAB_WIDTHS  = [900, 900, 700, 1200, 2880, 2780]  # total 9360
GRAMMAR_WIDTHS = [1400, 1800, 1360, 2500, 2300]     # total 9360

def build_vocab_table(vocab_list):
    headers = ["單字", "讀音", "詞性", "中文意思", "例句（日文）", "例句（中文解釋）"]
    header_row = "<w:tr>" + "".join(
        header_tc(h, w, "2E5FA3") for h, w in zip(headers, VOCAB_WIDTHS)
    ) + "</w:tr>"

    rows = ""
    for i, v in enumerate(vocab_list):
        fill = "EBF1FA" if i % 2 == 0 else "FFFFFF"
        # Build word cell with ruby
        word_ruby = ruby(v["word"], v["reading"], bold=True)
        sentence_xml = build_sentence_xml(v["sentence"], v["rubies"], v["plain_parts"])
        row = ("<w:tr>"
               + ruby_tc(word_ruby, VOCAB_WIDTHS[0], fill)
               + plain_tc(v["reading"], VOCAB_WIDTHS[1], fill)
               + plain_tc(v["pos"], VOCAB_WIDTHS[2], fill)
               + plain_tc(v["meaning"], VOCAB_WIDTHS[3], fill)
               + ruby_tc(sentence_xml, VOCAB_WIDTHS[4], fill)
               + plain_tc(v["translation"], VOCAB_WIDTHS[5], fill)
               + "</w:tr>")
        rows += row

    col_widths = "".join(f'<w:gridCol w:w="{w}"/>' for w in VOCAB_WIDTHS)
    return (f'<w:tbl>'
            f'<w:tblPr><w:tblW w:type="dxa" w:w="9360"/></w:tblPr>'
            f'<w:tblGrid>{col_widths}</w:tblGrid>'
            f'{header_row}{rows}'
            f'</w:tbl>')

def build_grammar_table(grammar_list):
    headers = ["語法型式", "意思／用法說明", "接續方式", "例句（日文）", "例句（中文解釋）"]
    header_row = "<w:tr>" + "".join(
        header_tc(h, w, "1A6640") for h, w in zip(headers, GRAMMAR_WIDTHS)
    ) + "</w:tr>"

    rows = ""
    for i, g in enumerate(grammar_list):
        fill = "E8F5EE" if i % 2 == 0 else "FFFFFF"
        sentence_xml = build_sentence_xml(g["sentence"], g["rubies"], g["plain_parts"])
        row = ("<w:tr>"
               + plain_tc(g["pattern"], GRAMMAR_WIDTHS[0], fill, bold=True)
               + multiline_tc(g["explanation"], GRAMMAR_WIDTHS[1], fill)
               + plain_tc(g["connection"], GRAMMAR_WIDTHS[2], fill)
               + ruby_tc(sentence_xml, GRAMMAR_WIDTHS[3], fill)
               + plain_tc(g["translation"], GRAMMAR_WIDTHS[4], fill)
               + "</w:tr>")
        rows += row

    col_widths = "".join(f'<w:gridCol w:w="{w}"/>' for w in GRAMMAR_WIDTHS)
    return (f'<w:tbl>'
            f'<w:tblPr><w:tblW w:type="dxa" w:w="9360"/></w:tblPr>'
            f'<w:tblGrid>{col_widths}</w:tblGrid>'
            f'{header_row}{rows}'
            f'</w:tbl>')

# ── Document XML template ─────────────────────────────────────────────────────

STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="240" w:after="200"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="280" w:after="160"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
</w:styles>'''

RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

ROOT_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

def build_document_xml(date_str, vocab_table_xml, grammar_table_xml):
    ns = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
          'mc:Ignorable="w14"')

    page_props = ('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'
                  '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" '
                  'w:header="720" w:footer="720" w:gutter="0"/>')

    title   = heading_para(f"日本語能力試驗 N2　單字・語法學習表", 1, "2E5FA3", 36, "center")
    sub     = subtitle_para(f"JLPT N2 Daily Study Sheet ／ {date_str}")
    vocab_h = heading_para("📘 N2 重要單字（3組）", 2, "2E5FA3", 28)
    gram_h  = heading_para("📗 N2 重要語法（3組）", 2, "1A6640", 28)
    foot    = footer_para("✦  加油！N2合格を目指して頑張りましょう！  ✦")

    body = (f"<w:body>"
            f"{title}{sub}"
            f"{vocab_h}{vocab_table_xml}"
            f"{spacer()}"
            f"{gram_h}{grammar_table_xml}"
            f"{foot}"
            f"<w:sectPr>{page_props}</w:sectPr>"
            f"</w:body>")

    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {ns}>{body}</w:document>'

# ── Main ──────────────────────────────────────────────────────────────────────

def generate(output_dir="."):
    # Taiwan time (UTC+8)
    tz = timezone(timedelta(hours=8))
    today = datetime.now(tz)
    date_str = today.strftime("%Y-%m-%d")
    filename = f"JLPT_N2_{date_str}.docx"

    # Seed with date for reproducibility within same day
    random.seed(date_str)
    vocab_pick   = random.sample(VOCAB_BANK, 3)
    grammar_pick = random.sample(GRAMMAR_BANK, 3)

    vocab_xml   = build_vocab_table(vocab_pick)
    grammar_xml = build_grammar_table(grammar_pick)
    doc_xml     = build_document_xml(date_str, vocab_xml, grammar_xml)

    # Pack into DOCX
    out_path = os.path.join(output_dir, filename)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml",              doc_xml)
        z.writestr("word/styles.xml",                STYLES_XML)
        z.writestr("word/_rels/document.xml.rels",   RELS_XML)
        z.writestr("[Content_Types].xml",            CONTENT_TYPES_XML)
        z.writestr("_rels/.rels",                    ROOT_RELS_XML)

    print(f"Generated: {out_path}")
    return filename

if __name__ == "__main__":
    generate(".")
