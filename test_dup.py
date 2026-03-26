"""
獨立測試：重複偵測候選生成
不需要 Gemini API，直接顯示哪些配對會被送給 AI 判斷
"""
import json

# 從圖片擷取的資料（單張工資表，當作同一頁處理）
# 格式：(編號, 姓名)
RAW = [
    ("5001", "肉清紫"),
    ("5099", "肉清紫"),  # 可能同名
    ("5242", "黃國治"),
    ("5516", "李友俊"),
    ("5865", "陳俊霖"),
    ("5389", "黃育林"),
    ("6036", "農育"),
    ("6073", "農育嘉"),
    ("6144", "潘雪岩"),
    ("6182", "費僚信"),
    ("6367", "廣方奐"),
    ("6747", "凌居務"),
    ("6782", "張逸餐"),
    ("6800", "金慶恩"),
    ("6927", "吳金益"),
    ("4111", "張登凱"),
    ("7148", "黃水財"),
    ("7159", "黃東為經"),
    ("7179", "黃天山"),
    ("7333", "玉謝程"),
    ("7223", "玉建松"),
    ("7382", "林家耀"),
    ("7371", "李家新"),
    ("4184", "莊家裁"),
    ("5110", "蔡振家"),
    ("5110", "林普德"),   # 🔴 同編號
    ("6097", "名烏"),
    ("5985", "呂冠良"),
    ("7844", "多拿"),
    ("7209", "余李"),
    ("4302", "古坤達"),
    ("8195", "沐明焦"),
    ("7480", "吳松"),
    ("5018", "陳文章"),
    ("5212", "許良澄"),
    ("7064", "劉祥梨"),
    ("9198", "余改逸"),
    ("6981", "賞柏翰"),
    ("5834", "陳冠杰"),
    ("9908", "湘南太"),
    ("6493", "古坤達"),
    ("5316", "張是祥"),
    ("5316", "為永俊"),   # 🔴 同編號
    ("5779", "黃仲凱"),
    ("8262", "蔡俊男"),
    ("1069", "徐冠廷"),
    ("6168", "廉廣浮"),
    ("7138", "黃新涼"),
    ("5281", "劉民泉"),
    ("5832", "劉民泉"),   # 🟠 同姓名
    ("9525", "許光雄"),
    ("7434", "劉佳"),
    ("5144", "李冠毅"),
    ("7487", "劉庭維"),
    ("4615", "陳敏雄"),
    ("4615", "陳敏雄"),   # 🔴 同編號+同姓名
    ("4489", "玩家興"),
    ("7469", "張家銘"),
    ("7984", "陳柏通"),
    ("4284", "陳孝先"),
    ("7311", "陳鎮鵬"),
    ("7322", "陳明"),
    ("8078", "張明文"),
    ("5381", "丁是網"),
    ("6317", "張逸餐"),
]

def run_candidate_detection(rows):
    """模擬 processor.py 的候選對生成邏輯（不呼叫 AI）"""
    exact_id   = []  # 🔴
    exact_name = []  # 🟠
    candidates = []  # 🟡 候選，待 AI

    n = len(rows)
    for i in range(n):
        for j in range(i + 1, n):
            id_a,   name_a = rows[i]
            id_b,   name_b = rows[j]

            # 🔴 編號相同 且 姓名相同 → 確定重複
            if id_a == id_b and name_a == name_b and len(name_a) >= 2:
                exact_id.append((rows[i], rows[j]))
                continue

            # 🟠 編號相同 但 姓名不同 → 直接標記
            if id_a == id_b and name_a != name_b and len(name_a) >= 2 and len(name_b) >= 2:
                exact_name.append((rows[i], rows[j]))
                continue

            # 🟡 近似編號候選
            if len(id_a) == len(id_b):
                diff = sum(x != y for x, y in zip(id_a, id_b))
                common_any = set(name_a) & set(name_b) - {'?', '', ' '}
                if 1 <= diff <= 2 and len(common_any) >= 1:
                    only_surname = (common_any == {name_a[0]}
                                    and name_a[0] == name_b[0]
                                    and diff == 2)
                    if not only_surname:
                        candidates.append({
                            'a': rows[i], 'b': rows[j],
                            'diff': diff,
                            'common': ''.join(sorted(common_any)),
                            'same_surname': name_a[0] == name_b[0] if name_a and name_b else False,
                        })

    return exact_id, exact_name, candidates


exact_id, exact_name, candidates = run_candidate_detection(RAW)

print("=" * 60)
print(f"🔴 同編號（直接標記，不需 AI）：{len(exact_id)} 對")
for a, b in exact_id:
    print(f"   {a[0]} {a[1]}  ↔  {b[0]} {b[1]}")

print()
print(f"🟠 同姓名（直接標記，不需 AI）：{len(exact_name)} 對")
for a, b in exact_name:
    print(f"   {a[0]} {a[1]}  ↔  {b[0]} {b[1]}")

print()
print(f"🟡 近似候選（送 AI 判斷）：{len(candidates)} 對")
print(f"   {'編號A':8} {'姓名A':8} {'編號B':8} {'姓名B':8} {'diff':4} {'共同字':6} {'同姓氏':5}")
print("   " + "-" * 55)
for c in candidates:
    a, b = c['a'], c['b']
    surname_mark = "✓" if c['same_surname'] else "✗"
    print(f"   {a[0]:8} {a[1]:8} {b[0]:8} {b[1]:8} {c['diff']:4}   {c['common']:6}  {surname_mark}")

print()
print(f"共 {len(candidates)} 對進入 AI 判斷（其中 {sum(1 for c in candidates if c['same_surname'])} 對姓氏相同，{sum(1 for c in candidates if not c['same_surname'])} 對姓氏不同但有共同字）")
